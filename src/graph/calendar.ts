/**
 * Calendar operations.
 *
 * Reads use `/me/calendarView` with a `Prefer: outlook.timezone` header so that
 * returned start/end times are rendered in the requested IANA zone (spec §7.3).
 * Writes send local wall-clock + IANA time zones (never a UTC `Z` with a named
 * zone) and set a Graph `transactionId` so retries do not create duplicates.
 */

import { createHash } from "node:crypto";
import { GraphError, ValidationError } from "../util/errors.js";
import { htmlToText } from "../util/html.js";
import {
  assertValidTimeZone,
  type GraphDateTime,
  preferTimeZoneHeader,
  toGraphDateTime,
  zonedWallClockToUtcIso,
} from "../util/time.js";
import type { GraphClient } from "./client.js";

const EVENT_SELECT = "id,subject,start,end,location,isAllDay,webLink";

/** The UTC instant (ms) of a local wall-clock time in its zone. */
function instant(t: { dateTime: string; timeZone: string }): number {
  return Date.parse(zonedWallClockToUtcIso(t.dateTime, t.timeZone));
}

/** Refuse a range or event that ends at or before it starts (#33). */
function assertEndAfterStart(startMs: number, endMs: number, message: string): void {
  if (endMs <= startMs) throw new ValidationError(message);
}

export interface ListEventsInput {
  startDateTime: string;
  endDateTime: string;
  timeZone: string;
  limit: number;
}

export interface EventTime {
  dateTime: string | null;
  timeZone: string | null;
}

export interface EventSummary {
  id: string;
  subject: string | null;
  start: EventTime | null;
  end: EventTime | null;
  location: string | null;
  isAllDay: boolean;
  webLink: string | null;
}

interface GraphEvent {
  id: string;
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  location?: { displayName?: string };
  isAllDay?: boolean;
  webLink?: string;
  body?: { contentType?: string; content?: string };
  isReminderOn?: boolean;
  reminderMinutesBeforeStart?: number;
  attendees?: unknown[];
  /** singleInstance | occurrence | exception | seriesMaster */
  type?: string;
  /** Version tag; sent back as `If-Match` so writes apply only to the checked version. */
  "@odata.etag"?: string;
}

function shapeTime(t?: { dateTime?: string; timeZone?: string }): EventTime | null {
  if (!t) return null;
  return { dateTime: t.dateTime ?? null, timeZone: t.timeZone ?? null };
}

function shapeEvent(e: GraphEvent): EventSummary {
  return {
    id: e.id,
    subject: e.subject ?? null,
    start: shapeTime(e.start),
    end: shapeTime(e.end),
    location: e.location?.displayName ?? null,
    isAllDay: Boolean(e.isAllDay),
    webLink: e.webLink ?? null,
  };
}

/**
 * List events in a date range via calendarView, rendered in `timeZone`.
 *
 * `startDateTime`/`endDateTime` are ISO 8601. A bare local wall-clock value is
 * interpreted in `timeZone` (so the queried window matches the zone the results
 * are rendered in — a naive "today in Brisbane" query won't silently miss or
 * include boundary events); a value carrying an offset/`Z` pins the instant
 * directly. The `Prefer` header controls the zone of the times in the response.
 */
export async function listEvents(
  graph: GraphClient,
  input: ListEventsInput,
): Promise<EventSummary[]> {
  assertValidTimeZone(input.timeZone);

  // Pin bare local bounds to the display zone so the query window and the
  // rendered results agree (see doc comment above). Unparseable bounds are
  // rejected here rather than passed to Graph.
  const startDateTime = zonedWallClockToUtcIso(input.startDateTime, input.timeZone);
  const endDateTime = zonedWallClockToUtcIso(input.endDateTime, input.timeZone);
  assertEndAfterStart(
    Date.parse(startDateTime),
    Date.parse(endDateTime),
    "endDateTime must be after startDateTime.",
  );

  const events = await graph.getAllPages<GraphEvent>(
    "/me/calendarView",
    {
      query: {
        startDateTime,
        endDateTime,
        $select: EVENT_SELECT,
        $orderby: "start/dateTime",
        $top: input.limit,
      },
      headers: { prefer: preferTimeZoneHeader(input.timeZone) },
    },
    { maxItems: input.limit },
  );

  return events.map(shapeEvent);
}

// ── Writes ───────────────────────────────────────────────────────────────────

export interface EventDateTime {
  /** Local wall-clock time (no offset/Z). */
  dateTime: string;
  /** IANA time zone. */
  timeZone: string;
}

export interface CreateEventInput {
  subject: string;
  start: EventDateTime;
  end: EventDateTime;
  location?: string;
  body?: string;
  isReminderOn?: boolean;
  reminderMinutesBeforeStart?: number;
  /** Idempotency key; generated if omitted. */
  transactionId?: string;
}

export interface UpdateEventInput {
  id: string;
  subject?: string;
  start?: EventDateTime;
  end?: EventDateTime;
  location?: string;
  body?: string;
  isReminderOn?: boolean;
  reminderMinutesBeforeStart?: number;
}

export interface EventWriteResult {
  id: string;
  subject: string | null;
  start: EventTime | null;
  end: EventTime | null;
  location: string | null;
  webLink: string | null;
  transactionId?: string;
}

function shapeWrite(e: GraphEvent, transactionId?: string): EventWriteResult {
  return {
    id: e.id,
    subject: e.subject ?? null,
    start: shapeTime(e.start),
    end: shapeTime(e.end),
    location: e.location?.displayName ?? null,
    webLink: e.webLink ?? null,
    ...(transactionId ? { transactionId } : {}),
  };
}

/**
 * The only event properties the connector ever writes (spec v2 §7.5).
 * Notably absent: `attendees` (and anything that triggers invitations or
 * meeting responses). Outlook emails invitations to attendees, so allowing
 * them would turn calendar access into a way to send mail, which the
 * connector deliberately can't do (no `Mail.Send`; see #57).
 */
export const WRITABLE_EVENT_FIELDS: ReadonlySet<string> = new Set([
  "subject",
  "start",
  "end",
  "location",
  "body",
  "isReminderOn",
  "reminderMinutesBeforeStart",
  "transactionId",
]);

/**
 * Defence in depth: refuse to send any property outside
 * {@link WRITABLE_EVENT_FIELDS}, so a future change can't quietly start
 * sending invitations.
 */
export function assertWritableEventPayload(payload: Record<string, unknown>): void {
  const extra = Object.keys(payload).filter((k) => !WRITABLE_EVENT_FIELDS.has(k));
  if (extra.length > 0) {
    throw new ValidationError(`Refusing to write event properties: ${extra.join(", ")}.`);
  }
}

/**
 * Derive a stable transactionId from the event's semantic fields, so an
 * accidental retry of the *same* create call dedupes at Graph instead of
 * creating a duplicate — which is what `idempotentHint: true` promises the
 * client. An explicit `input.transactionId` always takes precedence.
 */
function deterministicTransactionId(input: CreateEventInput): string {
  const material = [
    input.subject,
    input.start.dateTime,
    input.start.timeZone,
    input.end.dateTime,
    input.end.timeZone,
    input.location ?? "",
  ].join(String.fromCharCode(0));
  return createHash("sha256").update(material).digest("hex").slice(0, 36);
}

/**
 * Create a calendar event.
 *
 * Times are sent as local wall-clock + IANA time zone via {@link toGraphDateTime}
 * (which rejects offset/`Z` values), so the event lands at the intended local
 * time. When no `transactionId` is supplied a deterministic one is derived from
 * the event details, so a retried identical create does not duplicate.
 * `deriveId` is injectable for tests.
 */
export async function createEvent(
  graph: GraphClient,
  input: CreateEventInput,
  deriveId: (input: CreateEventInput) => string = deterministicTransactionId,
): Promise<EventWriteResult> {
  const transactionId = input.transactionId ?? deriveId(input);
  const start = toGraphDateTime(input.start.dateTime, input.start.timeZone);
  const end = toGraphDateTime(input.end.dateTime, input.end.timeZone);
  assertEndAfterStart(
    instant(start),
    instant(end),
    "The event's end must be after its start. Nothing was created.",
  );
  const payload: Record<string, unknown> = {
    subject: input.subject,
    start,
    end,
    transactionId,
    ...(input.location ? { location: { displayName: input.location } } : {}),
    ...(input.body ? { body: { contentType: "text", content: input.body } } : {}),
    ...(input.isReminderOn !== undefined ? { isReminderOn: input.isReminderOn } : {}),
    ...(input.reminderMinutesBeforeStart !== undefined
      ? { reminderMinutesBeforeStart: input.reminderMinutesBeforeStart }
      : {}),
  };
  assertWritableEventPayload(payload);
  const created = await graph.post<GraphEvent>("/me/events", payload);
  return shapeWrite(created, transactionId);
}

/** Fields `update_event` can change, in the order changes are reported. */
const UPDATABLE_FIELDS = [
  "subject",
  "start",
  "end",
  "location",
  "body",
  "isReminderOn",
  "reminderMinutesBeforeStart",
] as const;
type UpdatableField = (typeof UPDATABLE_FIELDS)[number];

/** Long bodies are cut to this many characters in a change report. */
const BODY_CHANGE_MAX_CHARS = 500;

/** One overwritten field: its value before and after the PATCH. */
export interface FieldChange {
  field: UpdatableField;
  before: unknown;
  after: unknown;
}

export interface EventUpdateResult {
  event: EventWriteResult;
  /** Before/after values of each field the update set (Outlook keeps no history). */
  changes: FieldChange[];
}

function bodyText(body: GraphEvent["body"]): string | null {
  if (!body?.content) return null;
  const text = (
    body.contentType?.toLowerCase() === "html" ? htmlToText(body.content) : body.content
  ).trim();
  return text.length > BODY_CHANGE_MAX_CHARS ? `${text.slice(0, BODY_CHANGE_MAX_CHARS)}…` : text;
}

function fieldValue(e: GraphEvent, field: UpdatableField): unknown {
  switch (field) {
    case "subject":
      return e.subject ?? null;
    case "start":
      return shapeTime(e.start);
    case "end":
      return shapeTime(e.end);
    case "location":
      return e.location?.displayName ?? null;
    case "body":
      return bodyText(e.body);
    case "isReminderOn":
      return e.isReminderOn ?? null;
    case "reminderMinutesBeforeStart":
      return e.reminderMinutesBeforeStart ?? null;
  }
}

/**
 * `If-Match` for the version that was checked, so a write can't hit a changed
 * event. Graph always returns `@odata.etag` for events. Without one, the write
 * would be unconditional and the checks above could be stale, so refuse.
 */
function ifMatch(e: GraphEvent, action: string): Record<string, string> {
  const etag = e["@odata.etag"]?.trim();
  if (!etag) {
    throw new ValidationError(
      `Can't ${action} this event safely: Microsoft Graph didn't return its version tag. ` +
        "Nothing was changed.",
    );
  }
  return { "if-match": etag };
}

/**
 * Graph answers 412 when `If-Match` no longer matches. That means the event
 * changed after it was checked, or (because the Graph client retries timed-out
 * requests) an earlier attempt of this very request already succeeded and its
 * response was lost. Both are possible, so the message mustn't claim nothing happened.
 */
async function onCheckedVersion<T>(write: () => Promise<T>): Promise<T> {
  try {
    return await write();
  } catch (e) {
    if (e instanceof GraphError && e.status === 412) {
      throw new ValidationError(
        "Microsoft Graph refused the write because the event no longer matches the version that " +
          "was checked. Either it was changed elsewhere, or an earlier attempt of this request " +
          "already went through. Check the event with list_events before trying again.",
      );
    }
    throw e;
  }
}

/**
 * Update an existing event with a partial patch (only provided fields).
 *
 * `update_event` overwrites in place with no undo, so the current values of the
 * fields being changed are read first and reported alongside the new ones.
 * Both requests use the same `Prefer: outlook.timezone` (the zone of the new
 * start/end), so before/after times are directly comparable. Reading first also
 * means a stale id fails before anything is written.
 *
 * Events with attendees are refused: changing a meeting you organised emails
 * the attendees an update, which is a send channel the connector forbids (spec
 * v2 §7.5, §9.2). The PATCH carries `If-Match` for the checked version, so an
 * event that changed in between (e.g. gained attendees) is not overwritten.
 */
export async function updateEvent(
  graph: GraphClient,
  input: UpdateEventInput,
): Promise<EventUpdateResult> {
  const patch: Record<string, unknown> = {};
  if (input.subject !== undefined) patch.subject = input.subject;
  if (input.start) patch.start = toGraphDateTime(input.start.dateTime, input.start.timeZone);
  if (input.end) patch.end = toGraphDateTime(input.end.dateTime, input.end.timeZone);
  if (input.location !== undefined) patch.location = { displayName: input.location };
  if (input.body !== undefined) patch.body = { contentType: "text", content: input.body };
  if (input.isReminderOn !== undefined) patch.isReminderOn = input.isReminderOn;
  if (input.reminderMinutesBeforeStart !== undefined) {
    patch.reminderMinutesBeforeStart = input.reminderMinutesBeforeStart;
  }
  if (Object.keys(patch).length === 0) {
    throw new ValidationError("update_event requires at least one field to change.");
  }
  assertWritableEventPayload(patch);

  const fields = UPDATABLE_FIELDS.filter((f) => f in patch);
  const zone = input.start?.timeZone ?? input.end?.timeZone;
  const headers = zone ? { prefer: preferTimeZoneHeader(zone) } : undefined;
  const path = `/me/events/${encodeURIComponent(input.id)}`;
  // Changing either side needs both, to check the result still ends after it starts.
  const select = new Set(["id", "attendees", ...fields, ...(zone ? ["start", "end"] : [])]);

  const before = await graph.get<GraphEvent>(path, {
    query: { $select: [...select].join(",") },
    headers,
  });
  if ((before.attendees?.length ?? 0) > 0) {
    throw new ValidationError(
      "This event has attendees, so changing it could email them an update. The connector " +
        "doesn't send mail; change it in Outlook instead. Nothing was changed.",
    );
  }
  if (zone) {
    // A side that isn't changing keeps its current value, which Graph rendered
    // in `zone` (the Prefer header), so both sides compare in the same zone.
    const current = (t?: { dateTime?: string }) =>
      t?.dateTime ? instant({ dateTime: t.dateTime, timeZone: zone }) : undefined;
    const startMs = input.start ? instant(patch.start as GraphDateTime) : current(before.start);
    const endMs = input.end ? instant(patch.end as GraphDateTime) : current(before.end);
    if (startMs !== undefined && endMs !== undefined) {
      assertEndAfterStart(
        startMs,
        endMs,
        "After this change the event would end at or before it starts. Nothing was changed.",
      );
    }
  }
  // Prepare the report's "before" side now: a failure here must not happen
  // after the write, where it would hide a change that was actually made.
  const beforeValues = fields.map((field) => fieldValue(before, field));
  const precondition = ifMatch(before, "update");

  const updated = await onCheckedVersion(() =>
    graph.patch<GraphEvent>(path, patch, { headers: { ...headers, ...precondition } }),
  );

  return {
    event: shapeWrite(updated),
    changes: fields.map((field, i) => ({
      field,
      before: beforeValues[i],
      after: fieldValue(updated, field),
    })),
  };
}

export interface CancelEventInput {
  id: string;
  /** Must match the event's title (case-insensitive): guards against a wrong id. */
  confirmSubject: string;
  /** IANA zone for the echoed start/end. */
  timeZone: string;
}

/**
 * Cancel (delete) an event: `DELETE /me/events/{id}`, which moves it to
 * Deleted Items (recoverable in Outlook; not `permanentDelete`).
 *
 * Guards (spec v2 §7.8), checked against the live event before deleting:
 *  - `confirmSubject` must match its title;
 *  - events with attendees are refused: deleting a meeting you organised
 *    emails cancellations, which is a send channel the connector forbids (#61);
 *  - anything recurring is refused: deleting a series master removes every
 *    occurrence, and deleting an occurrence removes it from its series rather
 *    than moving it to Deleted Items, so neither fits the soft-delete promise.
 * The DELETE carries `If-Match` for the checked version, so an event that
 * changed in between (e.g. gained attendees) is not deleted.
 * Returns the removed event so it can be echoed to the user.
 */
export async function cancelEvent(
  graph: GraphClient,
  input: CancelEventInput,
): Promise<EventWriteResult> {
  assertValidTimeZone(input.timeZone);
  const path = `/me/events/${encodeURIComponent(input.id)}`;
  const event = await graph.get<GraphEvent>(path, {
    query: { $select: "id,subject,start,end,location,webLink,attendees,type" },
    headers: { prefer: preferTimeZoneHeader(input.timeZone) },
  });

  const title = (event.subject ?? "").trim();
  if (title.toLowerCase() !== input.confirmSubject.trim().toLowerCase()) {
    throw new ValidationError(
      `confirmSubject "${input.confirmSubject}" does not match the event's title "${title}". ` +
        "Nothing was deleted. Check the id with list_events and confirm with the user.",
    );
  }
  if ((event.attendees?.length ?? 0) > 0) {
    throw new ValidationError(
      "This event has attendees, so deleting it could email them a cancellation. The connector " +
        "doesn't send mail; cancel it in Outlook instead. Nothing was deleted.",
    );
  }
  if (event.type === "seriesMaster") {
    throw new ValidationError(
      "This is a whole recurring series; deleting it would remove every occurrence. Manage the " +
        "series in Outlook instead. Nothing was deleted.",
    );
  }
  if (event.type === "occurrence" || event.type === "exception") {
    throw new ValidationError(
      "This is one occurrence of a recurring series; deleting it removes it from the series " +
        "instead of moving it to Deleted Items. Cancel it in Outlook instead. Nothing was deleted.",
    );
  }

  const precondition = ifMatch(event, "delete");
  await onCheckedVersion(() => graph.delete(path, { headers: precondition }));
  return shapeWrite(event);
}
