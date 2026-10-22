/**
 * Mail read operations (Mail.Read).
 *
 * Graph does not allow `$search` to be combined with `$filter`/`$orderby`, so
 * the query builder chooses ONE mode:
 *  - any text intent (query / from / subjectContains) -> `$search` (KQL); date
 *    bounds, if also given, are applied client-side to the returned page;
 *  - otherwise -> `$filter` on receivedDateTime with `$orderby ... desc`.
 * This keeps every request Graph-valid.
 */

import { z } from "zod";
import { htmlToText } from "../util/html.js";
import { GraphError, ValidationError } from "../util/errors.js";
import type { GraphClient, ODataPage } from "./client.js";
import { CursorBase, decodeCursor, encodeCursor } from "./paging.js";

const MESSAGE_SELECT =
  "id,subject,from,receivedDateTime,webLink,bodyPreview,hasAttachments,parentFolderId";

export interface ListMessagesInput {
  query?: string;
  from?: string;
  subjectContains?: string;
  receivedAfter?: string;
  receivedBefore?: string;
  limit: number;
  pageToken?: string;
  /**
   * Only this folder: a well-known name (see {@link WELL_KNOWN_FOLDERS}) or a
   * folder id from list_mail_folders. All folders if omitted (#53).
   */
  folder?: string;
  /** Without a folder, Junk Email and Deleted Items are left out unless this is set. */
  includeJunkAndDeleted?: boolean;
}

/**
 * Folder names Graph accepts in place of an id. `junkemail` and `deleteditems`
 * are the ones left out by default; the others label results.
 */
export const WELL_KNOWN_FOLDERS = [
  "inbox",
  "junkemail",
  "deleteditems",
  "archive",
  "sentitems",
  "drafts",
] as const;
export type WellKnownFolder = (typeof WELL_KNOWN_FOLDERS)[number];
/** Left out of a listing across folders unless asked for (#53). */
const EXCLUDED_BY_DEFAULT: readonly WellKnownFolder[] = ["junkemail", "deleteditems"];

export interface MessageSummary {
  id: string;
  subject: string | null;
  from: { name?: string; address?: string } | null;
  receivedDateTime: string | null;
  webLink: string | null;
  bodyPreview: string | null;
  hasAttachments: boolean;
  /** The message's folder id (`parentFolderId`). */
  folderId: string | null;
  /** Its well-known name when it's one of those (e.g. "junkemail"), else null. */
  folder: WellKnownFolder | null;
}

export interface ListMessagesResult {
  messages: MessageSummary[];
  nextPageToken?: string;
  /** Messages on this page left out because they're in Junk Email or Deleted Items. */
  excludedCount: number;
}

/**
 * Folder lookups in flight at once. Outlook allows four concurrent requests per
 * mailbox, and a listing has its own request running alongside these.
 */
const FOLDER_LOOKUPS_AT_ONCE = 3;
/** Most subfolders of Junk Email and Deleted Items a listing walks to leave them out. */
export const MAX_EXCLUDED_SUBFOLDERS = 200;

/** `fn` over `items`, at most `limit` at a time, results in order. */
async function mapLimit<T, R>(items: readonly T[], limit: number, fn: (item: T) => Promise<R>) {
  const out: R[] = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i]!);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return out;
}

interface FolderRef {
  id: string;
  childFolderCount: number;
}
interface KnownFolder extends FolderRef {
  name: WellKnownFolder;
}

/**
 * The well-known folders the mailbox has, so results can be labelled and Junk /
 * Deleted left out. Looked up per call; a folder the mailbox doesn't have (e.g.
 * no Archive) is simply absent.
 */
async function wellKnownFolders(graph: GraphClient): Promise<KnownFolder[]> {
  const found = await mapLimit(WELL_KNOWN_FOLDERS, FOLDER_LOOKUPS_AT_ONCE, async (name) => {
    try {
      const f = await graph.get<{ id?: string; childFolderCount?: number }>(
        `/me/mailFolders/${name}`,
        { query: { $select: "id,childFolderCount" } },
      );
      return f.id ? { id: f.id, name, childFolderCount: f.childFolderCount ?? 0 } : undefined;
    } catch (e) {
      if (e instanceof GraphError && e.status === 404) return undefined;
      throw e;
    }
  });
  return found.filter((f): f is KnownFolder => f !== undefined);
}

function folderLabels(known: KnownFolder[]): Map<string, WellKnownFolder> {
  return new Map(known.map((f) => [f.id, f.name]));
}

const TOO_MANY_SUBFOLDERS =
  "Junk Email and Deleted Items hold too many subfolders to leave them out reliably. " +
  "Pass a folder, or includeJunkAndDeleted: true.";

/** A folder's direct subfolders; none if it has gone meanwhile. */
async function childFolders(graph: GraphClient, id: string): Promise<FolderRef[]> {
  try {
    const page = await graph.get<ODataPage<{ id: string; childFolderCount?: number }>>(
      `/me/mailFolders/${encodeURIComponent(id)}/childFolders`,
      {
        query: {
          $select: "id,childFolderCount",
          $top: MAX_EXCLUDED_SUBFOLDERS,
          // Hidden folders too: the listing returns their messages all the same.
          includeHiddenFolders: "true",
        },
      },
    );
    if (page["@odata.nextLink"]) throw new ValidationError(TOO_MANY_SUBFOLDERS);
    return (page.value ?? []).map((f) => ({ id: f.id, childFolderCount: f.childFolderCount ?? 0 }));
  } catch (e) {
    if (e instanceof GraphError && e.status === 404) return [];
    throw e;
  }
}

/**
 * Junk Email and Deleted Items, and every folder beneath them: a deleted folder
 * moves under Deleted Items with its messages still filed in it, so the two
 * roots alone would let them through. Refuses, rather than leak, past
 * {@link MAX_EXCLUDED_SUBFOLDERS}, counted as each response arrives so a huge
 * tree is refused early rather than walked in full.
 */
function excludedFolderIds(graph: GraphClient, known: KnownFolder[]): Promise<Set<string>> {
  const roots = known.filter((f) => EXCLUDED_BY_DEFAULT.includes(f.name));
  const ids = new Set(roots.map((f) => f.id));
  const queue = roots.filter((f) => f.childFolderCount > 0).map((f) => f.id);
  return new Promise((resolve, reject) => {
    let running = 0;
    let failed = false;
    const fail = (e: unknown) => {
      failed = true;
      reject(e);
    };
    // Start lookups up to the limit; finish once nothing is queued or running.
    const pump = () => {
      if (failed) return;
      if (queue.length === 0 && running === 0) {
        resolve(ids);
        return;
      }
      while (running < FOLDER_LOOKUPS_AT_ONCE && queue.length > 0) {
        running++;
        childFolders(graph, queue.shift()!).then((children) => {
          running--;
          if (failed) return;
          for (const c of children) {
            if (ids.has(c.id)) continue;
            ids.add(c.id);
            if (ids.size - roots.length > MAX_EXCLUDED_SUBFOLDERS) {
              fail(new ValidationError(TOO_MANY_SUBFOLDERS));
              return;
            }
            if (c.childFolderCount > 0) queue.push(c.id);
          }
          pump();
        }, fail);
      }
    };
    pump();
  });
}

interface GraphMessage {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  receivedDateTime?: string;
  webLink?: string;
  bodyPreview?: string;
  hasAttachments?: boolean;
  body?: { contentType?: string; content?: string };
  parentFolderId?: string;
}

function toInstant(iso: string, label: string): string {
  const t = Date.parse(iso);
  if (Number.isNaN(t)) {
    throw new ValidationError(`${label} must be an ISO 8601 date/time; got "${iso}".`);
  }
  return new Date(t).toISOString();
}

export interface BuiltMessagesQuery {
  /** Which Graph mode the query uses; drives the eventual-consistency header. */
  mode: "search" | "filter";
  query: Record<string, string | number>;
  clientDateFilter?: { afterMs?: number; beforeMs?: number };
}

/** Build the Graph query for /me/messages from the tool input. */
export function buildMessagesQuery(input: ListMessagesInput): BuiltMessagesQuery {
  const q: Record<string, string | number> = {
    $select: MESSAGE_SELECT,
    $top: input.limit,
  };

  const terms: string[] = [];
  if (input.query) terms.push(input.query.trim());
  if (input.from) terms.push(`from:${input.from.trim()}`);
  if (input.subjectContains) terms.push(`subject:${input.subjectContains.trim()}`);

  if (terms.length > 0) {
    // $search mode — cannot combine with $filter/$orderby. Strip embedded
    // double-quotes so a term can't break out of the quoted KQL phrase.
    q.$search = `"${terms.map((t) => t.replace(/"/g, "")).join(" ")}"`;
    const clientDateFilter: { afterMs?: number; beforeMs?: number } = {};
    if (input.receivedAfter)
      clientDateFilter.afterMs = Date.parse(toInstant(input.receivedAfter, "receivedAfter"));
    if (input.receivedBefore)
      clientDateFilter.beforeMs = Date.parse(toInstant(input.receivedBefore, "receivedBefore"));
    return {
      mode: "search",
      query: q,
      ...(clientDateFilter.afterMs !== undefined || clientDateFilter.beforeMs !== undefined
        ? { clientDateFilter }
        : {}),
    };
  }

  // $filter mode.
  const filters: string[] = [];
  if (input.receivedAfter)
    filters.push(`receivedDateTime ge ${toInstant(input.receivedAfter, "receivedAfter")}`);
  if (input.receivedBefore)
    filters.push(`receivedDateTime le ${toInstant(input.receivedBefore, "receivedBefore")}`);
  if (filters.length > 0) q.$filter = filters.join(" and ");
  q.$orderby = "receivedDateTime desc";
  return { mode: "filter", query: q };
}

function shapeSummary(m: GraphMessage, folders: Map<string, WellKnownFolder>): MessageSummary {
  const ea = m.from?.emailAddress;
  return {
    id: m.id,
    subject: m.subject ?? null,
    from: ea ? { name: ea.name, address: ea.address } : null,
    receivedDateTime: m.receivedDateTime ?? null,
    webLink: m.webLink ?? null,
    bodyPreview: m.bodyPreview ?? null,
    hasAttachments: Boolean(m.hasAttachments),
    folderId: m.parentFolderId ?? null,
    folder: (m.parentFolderId && folders.get(m.parentFolderId)) || null,
  };
}

// ── Paging cursor ────────────────────────────────────────────────────────────
//
// See ./paging.ts. Besides the next page, a list_messages cursor holds the query
// mode (for the eventual-consistency header), any client-side date bounds, and
// whether Junk / Deleted are left out, so later pages behave like the first
// (#31, #53).

const CursorSchema = CursorBase.extend({
  mode: z.enum(["search", "filter"]),
  afterMs: z.number().finite().optional(),
  beforeMs: z.number().finite().optional(),
  excludeJunkAndDeleted: z.boolean().optional(),
});
type Cursor = z.infer<typeof CursorSchema>;

function withinDateFilter(m: GraphMessage, f?: { afterMs?: number; beforeMs?: number }): boolean {
  if (!f || !m.receivedDateTime) return true;
  const t = Date.parse(m.receivedDateTime);
  if (Number.isNaN(t)) return true;
  if (f.afterMs !== undefined && t < f.afterMs) return false;
  if (f.beforeMs !== undefined && t > f.beforeMs) return false;
  return true;
}

/** List / search messages. */
export async function listMessages(
  graph: GraphClient,
  input: ListMessagesInput,
): Promise<ListMessagesResult> {
  // With a pageToken, the cursor alone decides the request; other inputs are ignored.
  let mode: Cursor["mode"];
  let path: string;
  let query: Record<string, string | number> | undefined;
  let clientDateFilter: { afterMs?: number; beforeMs?: number } | undefined;
  let excludeJunkAndDeleted: boolean;

  if (input.pageToken) {
    const cursor = decodeCursor(input.pageToken, CursorSchema, "list_messages");
    mode = cursor.mode;
    path = cursor.next;
    if (cursor.afterMs !== undefined || cursor.beforeMs !== undefined) {
      clientDateFilter = { afterMs: cursor.afterMs, beforeMs: cursor.beforeMs };
    }
    excludeJunkAndDeleted = cursor.excludeJunkAndDeleted ?? false;
  } else {
    const built = buildMessagesQuery(input);
    mode = built.mode;
    // A folder narrows the listing to it; otherwise it spans every folder, less
    // Junk and Deleted unless asked (phishing in Junk must not pass for real
    // mail, #53). Graph can't combine $search with a $filter on the folder, so
    // that exclusion is applied to each page as it comes.
    const folder = input.folder?.trim();
    // "." or ".." would be resolved away in the URL, listing the whole mailbox
    // with nothing left out.
    if (folder !== undefined && (folder === "" || folder === "." || folder === "..")) {
      throw new ValidationError(
        "folder must be a well-known folder name or a folder id from list_mail_folders.",
      );
    }
    path = folder ? `/me/mailFolders/${encodeURIComponent(folder)}/messages` : "/me/messages";
    query = built.query;
    clientDateFilter = built.clientDateFilter;
    excludeJunkAndDeleted = !folder && !input.includeJunkAndDeleted;
  }

  const [page, { folders, excluded }] = await Promise.all([
    graph
      .request<ODataPage<GraphMessage>>("GET", path, {
        query,
        // $search on messages benefits from eventual consistency.
        headers: mode === "search" ? { consistencylevel: "eventual" } : undefined,
      })
      .then((r) => r.data),
    wellKnownFolders(graph).then(async (known) => ({
      folders: folderLabels(known),
      excluded: excludeJunkAndDeleted ? await excludedFolderIds(graph, known) : new Set<string>(),
    })),
  ]);

  const inDates = (page.value ?? []).filter((m) => withinDateFilter(m, clientDateFilter));
  const kept = inDates.filter((m) => !(m.parentFolderId && excluded.has(m.parentFolderId)));
  const nextPageToken = encodeCursor(page["@odata.nextLink"], {
    mode,
    ...clientDateFilter,
    ...(excludeJunkAndDeleted ? { excludeJunkAndDeleted } : {}),
  });

  return {
    messages: kept.map((m) => shapeSummary(m, folders)),
    excludedCount: inDates.length - kept.length,
    ...(nextPageToken ? { nextPageToken } : {}),
  };
}

export interface MailFolder {
  id: string;
  name: string | null;
  /** Its well-known name (e.g. "junkemail"), when it's one of those. */
  wellKnownName: WellKnownFolder | null;
  parentFolderId: string | null;
  childFolderCount: number;
  unreadItemCount: number;
  totalItemCount: number;
}

/** Most folders list_mail_folders returns. */
export const MAX_MAIL_FOLDERS = 100;

/** The mailbox's top-level folders (`GET /me/mailFolders`), for list_messages' `folder` (#53). */
export async function listMailFolders(graph: GraphClient): Promise<MailFolder[]> {
  const [raw, known] = await Promise.all([
    graph.getAllPages<{
      id: string;
      displayName?: string;
      parentFolderId?: string;
      childFolderCount?: number;
      unreadItemCount?: number;
      totalItemCount?: number;
    }>(
      "/me/mailFolders",
      {
        query: {
          $select: "id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount",
          $top: MAX_MAIL_FOLDERS,
        },
      },
      { maxItems: MAX_MAIL_FOLDERS },
    ),
    wellKnownFolders(graph).then(folderLabels),
  ]);
  return raw.map((f) => ({
    id: f.id,
    name: f.displayName ?? null,
    wellKnownName: known.get(f.id) ?? null,
    parentFolderId: f.parentFolderId ?? null,
    childFolderCount: f.childFolderCount ?? 0,
    unreadItemCount: f.unreadItemCount ?? 0,
    totalItemCount: f.totalItemCount ?? 0,
  }));
}

/** Default cap on a returned body, so one newsletter can't flood the model's context (#38). */
export const DEFAULT_MAX_BODY_CHARS = 20_000;

export interface GetMessageInput {
  id: string;
  format: "text" | "html";
  /** Most body characters to return (default {@link DEFAULT_MAX_BODY_CHARS}). */
  maxChars?: number;
  /** Where in the (converted) body to start, to read past a previous cut. */
  startChar?: number;
}

export interface MessageBody {
  contentType: "text" | "html";
  content: string;
  /** True when characters after this part were left out. */
  truncated: boolean;
  /** Length of the whole (converted) body. */
  totalChars: number;
  /** Where this part starts; present when reading past the beginning. */
  startChar?: number;
  /** Pass as `startChar` to read the next part; present when truncated. */
  nextStartChar?: number;
}

/**
 * The part of `content` from `startChar`, at most `maxChars` long. Cuts never
 * split a surrogate pair, so emoji and other astral characters stay whole.
 */
export function sliceBody(
  content: string,
  contentType: "text" | "html",
  maxChars = DEFAULT_MAX_BODY_CHARS,
  startChar = 0,
): MessageBody {
  // True when `i` falls between the two units of a real pair. A lone surrogate
  // (malformed mail) isn't a pair, so it never moves a cut.
  const splitsPair = (i: number) => {
    if (i <= 0 || i >= content.length) return false;
    const hi = content.charCodeAt(i - 1);
    const lo = content.charCodeAt(i);
    return hi >= 0xd800 && hi <= 0xdbff && lo >= 0xdc00 && lo <= 0xdfff;
  };
  let start = Math.min(Math.max(0, startChar), content.length);
  if (splitsPair(start)) start -= 1; // don't begin halfway through a pair
  let end = Math.min(start + maxChars, content.length);
  // Nor end halfway: back off, or take the whole pair if that would leave nothing.
  if (splitsPair(end)) end = end - 1 > start ? end - 1 : end + 1;
  const truncated = end < content.length;
  return {
    contentType,
    content: content.slice(start, end),
    truncated,
    totalChars: content.length,
    ...(start > 0 ? { startChar: start } : {}),
    ...(truncated ? { nextStartChar: end } : {}),
  };
}

export interface AttachmentMeta {
  id?: string;
  name: string | null;
  contentType: string | null;
  size: number | null;
  isInline: boolean;
}

export interface MessageDetail {
  id: string;
  subject: string | null;
  from: { name?: string; address?: string } | null;
  receivedDateTime: string | null;
  webLink: string | null;
  body: MessageBody;
  hasAttachments: boolean;
  attachments: AttachmentMeta[];
  /** The message's folder id, and its well-known name (e.g. "junkemail") when it's one of those. */
  folderId: string | null;
  folder: WellKnownFolder | null;
}

interface GraphAttachment {
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
}

/**
 * Read a single message, converting HTML to text unless html is requested. The
 * body is capped (see {@link sliceBody}); `startChar` reads further parts.
 */
export async function getMessage(
  graph: GraphClient,
  input: GetMessageInput,
): Promise<MessageDetail> {
  const id = encodeURIComponent(input.id);
  const [msg, folders] = await Promise.all([
    graph.get<GraphMessage>(`/me/messages/${id}`, {
      query: {
        $select: "id,subject,from,receivedDateTime,webLink,body,hasAttachments,parentFolderId",
      },
    }),
    wellKnownFolders(graph).then(folderLabels),
  ]);

  const rawContent = msg.body?.content ?? "";
  const rawType = (msg.body?.contentType ?? "text").toLowerCase();
  let content = rawContent;
  if (input.format === "text" && rawType === "html") {
    content = htmlToText(rawContent);
  }
  // What the content really is (#35): HTML only if it was HTML and HTML was
  // asked for. A plain-text message stays text even when format=html.
  const contentType = rawType === "html" && input.format === "html" ? "html" : "text";

  let attachments: AttachmentMeta[] = [];
  if (msg.hasAttachments) {
    const page = await graph.get<ODataPage<GraphAttachment>>(`/me/messages/${id}/attachments`, {
      query: { $select: "id,name,contentType,size,isInline" },
    });
    attachments = (page.value ?? []).map((a) => ({
      id: a.id,
      name: a.name ?? null,
      contentType: a.contentType ?? null,
      size: a.size ?? null,
      isInline: Boolean(a.isInline),
    }));
  }

  const ea = msg.from?.emailAddress;
  return {
    id: msg.id,
    subject: msg.subject ?? null,
    from: ea ? { name: ea.name, address: ea.address } : null,
    receivedDateTime: msg.receivedDateTime ?? null,
    webLink: msg.webLink ?? null,
    body: sliceBody(content, contentType, input.maxChars, input.startChar),
    hasAttachments: Boolean(msg.hasAttachments),
    attachments,
    folderId: msg.parentFolderId ?? null,
    folder: (msg.parentFolderId && folders.get(msg.parentFolderId)) || null,
  };
}
