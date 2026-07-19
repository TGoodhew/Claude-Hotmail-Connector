/**
 * Calendar operations.
 *
 * Reads use `/me/calendarView` with a `Prefer: outlook.timezone` header so that
 * returned start/end times are rendered in the requested IANA zone (spec §7.3).
 * Writes send local wall-clock + IANA time zones (never a UTC `Z` with a named
 * zone) and set a Graph `transactionId` so retries do not create duplicates.
 */

import { randomUUID } from "node:crypto";
import { ValidationError } from "../util/errors.js";
import { assertValidTimeZone, preferTimeZoneHeader, toGraphDateTime } from "../util/time.js";
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
 * `startDateTime`/`endDateTime` should be ISO 8601. Include a timezone offset
 * (or `Z`) to pin the window precisely; without one Graph treats the window as
 * UTC. The `Prefer` header controls the zone of the times in the response.
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

  const events = await graph.getAllPages<GraphEvent>(
    "/me/calendarView",
    {
      query: {
        startDateTime: input.startDateTime,
        endDateTime: input.endDateTime,
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
 * Create a calendar event.
 *
 * Times are sent as local wall-clock + IANA time zone via {@link toGraphDateTime}
 * (which rejects offset/`Z` values), so the event lands at the intended local
 * time. A `transactionId` is set for idempotency — repeated calls with the same
 * id do not create duplicates. `generateId` is injectable for tests.
 */
export async function createEvent(
  graph: GraphClient,
  input: CreateEventInput,
  generateId: () => string = randomUUID,
): Promise<EventWriteResult> {
  const transactionId = input.transactionId ?? generateId();
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
  const created = await graph.post<GraphEvent>("/me/events", payload);
  return shapeWrite(created, transactionId);
}

/** Update an existing event with a partial patch (only provided fields). */
export async function updateEvent(
  graph: GraphClient,
  input: UpdateEventInput,
): Promise<EventWriteResult> {
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
  const updated = await graph.patch<GraphEvent>(
    `/me/events/${encodeURIComponent(input.id)}`,
    patch,
  );
  return shapeWrite(updated);
}
