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
import { GRAPH_BASE_URL } from "../config.js";
import { htmlToText } from "../util/html.js";
import { ValidationError } from "../util/errors.js";
import type { GraphClient, ODataPage } from "./client.js";

const MESSAGE_SELECT = "id,subject,from,receivedDateTime,webLink,bodyPreview,hasAttachments";

export interface ListMessagesInput {
  query?: string;
  from?: string;
  subjectContains?: string;
  receivedAfter?: string;
  receivedBefore?: string;
  limit: number;
  pageToken?: string;
}

export interface MessageSummary {
  id: string;
  subject: string | null;
  from: { name?: string; address?: string } | null;
  receivedDateTime: string | null;
  webLink: string | null;
  bodyPreview: string | null;
  hasAttachments: boolean;
}

export interface ListMessagesResult {
  messages: MessageSummary[];
  nextPageToken?: string;
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

function shapeSummary(m: GraphMessage): MessageSummary {
  const ea = m.from?.emailAddress;
  return {
    id: m.id,
    subject: m.subject ?? null,
    from: ea ? { name: ea.name, address: ea.address } : null,
    receivedDateTime: m.receivedDateTime ?? null,
    webLink: m.webLink ?? null,
    bodyPreview: m.bodyPreview ?? null,
    hasAttachments: Boolean(m.hasAttachments),
  };
}

// ── Paging cursor ────────────────────────────────────────────────────────────
//
// `nextPageToken` is an opaque cursor, never a raw URL (spec v2 §7.2, #30): it
// holds the next page's path *relative to the Graph base* plus what later pages
// need to behave like the first, i.e. the query mode (for the
// eventual-consistency header) and any client-side date bounds (#31). A
// relative path can only ever be resolved against Graph, and the Graph client
// rejects absolute URLs outside it besides.

const CursorSchema = z.object({
  v: z.literal(1),
  next: z
    .string()
    .min(1)
    .refine((p) => /^[A-Za-z]/.test(p) && !p.includes("://") && !p.includes("\\"), {
      message: "must be a relative Graph path",
    }),
  mode: z.enum(["search", "filter"]),
  afterMs: z.number().finite().optional(),
  beforeMs: z.number().finite().optional(),
});
type Cursor = z.infer<typeof CursorSchema>;

const GRAPH_PREFIX = GRAPH_BASE_URL.replace(/\/?$/, "/");

function encodeCursor(
  nextLink: string | undefined,
  mode: Cursor["mode"],
  dates?: { afterMs?: number; beforeMs?: number },
): string | undefined {
  // Graph's own nextLinks live under the base; anything else is dropped, never followed.
  if (!nextLink?.startsWith(GRAPH_PREFIX)) return undefined;
  const cursor: Cursor = { v: 1, next: nextLink.slice(GRAPH_PREFIX.length), mode, ...dates };
  return Buffer.from(JSON.stringify(cursor)).toString("base64url");
}

function decodeCursor(token: string): Cursor {
  let raw: unknown;
  try {
    raw = JSON.parse(Buffer.from(token, "base64url").toString("utf8"));
  } catch {
    raw = undefined;
  }
  const parsed = CursorSchema.safeParse(raw);
  if (!parsed.success) {
    throw new ValidationError(
      "pageToken is not a valid cursor. Pass the nextPageToken from a previous list_messages call " +
        "unchanged (it is not a URL).",
    );
  }
  return parsed.data;
}

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

  if (input.pageToken) {
    const cursor = decodeCursor(input.pageToken);
    mode = cursor.mode;
    path = cursor.next;
    if (cursor.afterMs !== undefined || cursor.beforeMs !== undefined) {
      clientDateFilter = { afterMs: cursor.afterMs, beforeMs: cursor.beforeMs };
    }
  } else {
    const built = buildMessagesQuery(input);
    mode = built.mode;
    path = "/me/messages";
    query = built.query;
    clientDateFilter = built.clientDateFilter;
  }

  const page = (
    await graph.request<ODataPage<GraphMessage>>("GET", path, {
      query,
      // $search on messages benefits from eventual consistency.
      headers: mode === "search" ? { consistencylevel: "eventual" } : undefined,
    })
  ).data;

  const messages = (page.value ?? [])
    .filter((m) => withinDateFilter(m, clientDateFilter))
    .map(shapeSummary);
  const nextPageToken = encodeCursor(page["@odata.nextLink"], mode, clientDateFilter);

  return { messages, ...(nextPageToken ? { nextPageToken } : {}) };
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
  const msg = await graph.get<GraphMessage>(`/me/messages/${id}`, {
    query: { $select: "id,subject,from,receivedDateTime,webLink,body,hasAttachments" },
  });

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
  };
}
