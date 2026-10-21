/**
 * Calendar operations.
 *
 * Reads use `/me/calendarView` with a `Prefer: outlook.timezone` header so that
 * returned start/end times are rendered in the requested IANA zone (spec §7.3).
 * Writes send local wall-clock + IANA time zones (never a UTC `Z` with a named
 * zone) and set a Graph `transactionId` so retries do not create duplicates.
 */

import { createHash } from "node:crypto";
import { ValidationError } from "../util/errors.js";
import { htmlToText } from "../util/html.js";
import {
  assertValidTimeZone,
  preferTimeZoneHeader,
  toGraphDateTime,
  zonedWallClockToUtcIso,
} from "../util/time.js";
import type { GraphClient } from "./client.js";

const EVENT_SELECT = "id,subject,start,end,location,isAllDay,webLink";

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
  if (Number.isNaN(Date.parse(input.startDateTime))) {
    throw new ValidationError(`startDateTime must be ISO 8601; got "${input.startDateTime}".`);
  }
  if (Number.isNaN(Date.parse(input.endDateTime))) {
    throw new ValidationError(`endDateTime must be ISO 8601; got "${input.endDateTime}".`);
  }

  // Pin bare local bounds to the display zone so the query window and the
  // rendered results agree (see doc comment above).
  const startDateTime = zonedWallClockToUtcIso(input.startDateTime, input.timeZone);
  const endDateTime = zonedWallClockToUtcIso(input.endDateTime, input.timeZone);

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
  const payload: Record<string, unknown> = {
    subject: input.subject,
    start: toGraphDateTime(input.start.dateTime, input.start.timeZone),
    end: toGraphDateTime(input.end.dateTime, input.end.timeZone),
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
 * Update an existing event with a partial patch (only provided fields).
 *
 * `update_event` overwrites in place with no undo, so the current values of the
 * fields being changed are read first and reported alongside the new ones.
 * Both requests use the same `Prefer: outlook.timezone` (the zone of the new
 * start/end), so before/after times are directly comparable. Reading first also
 * means a stale id fails before anything is written.
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

  const before = await graph.get<GraphEvent>(path, {
    query: { $select: ["id", ...fields].join(",") },
    headers,
  });
  const updated = await graph.patch<GraphEvent>(path, patch, { headers });

  return {
    event: shapeWrite(updated),
    changes: fields.map((field) => ({
      field,
      before: fieldValue(before, field),
      after: fieldValue(updated, field),
    })),
  };
}
