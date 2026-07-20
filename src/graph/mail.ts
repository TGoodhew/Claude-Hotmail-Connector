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
  // A pageToken is a full @odata.nextLink URL — follow it verbatim.
  let page: ODataPage<GraphMessage>;
  let clientDateFilter: { afterMs?: number; beforeMs?: number } | undefined;

  if (input.pageToken) {
    page = (await graph.request<ODataPage<GraphMessage>>("GET", input.pageToken)).data;
  } else {
    const built = buildMessagesQuery(input);
    clientDateFilter = built.clientDateFilter;
    page = (
      await graph.request<ODataPage<GraphMessage>>("GET", "/me/messages", {
        query: built.query,
        // $search on messages benefits from eventual consistency.
        headers: built.mode === "search" ? { consistencylevel: "eventual" } : undefined,
      })
    ).data;
  }

  const messages = (page.value ?? [])
    .filter((m) => withinDateFilter(m, clientDateFilter))
    .map(shapeSummary);

  return {
    messages,
    ...(page["@odata.nextLink"] ? { nextPageToken: page["@odata.nextLink"] } : {}),
  };
}

export interface GetMessageInput {
  id: string;
  format: "text" | "html";
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
  body: { contentType: "text" | "html"; content: string };
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

/** Read a single message, converting HTML to text unless html is requested. */
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
    body: { contentType: input.format, content },
    hasAttachments: Boolean(msg.hasAttachments),
    attachments,
  };
}
