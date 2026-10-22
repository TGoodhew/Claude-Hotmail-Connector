/**
 * Calendar operations.
 *
 * Reads use `/me/calendarView` with a `Prefer: outlook.timezone` header so that
 * returned start/end times are rendered in the requested IANA zone (spec §7.3).
 * Writes send local wall-clock + IANA time zones (never a UTC `Z` with a named
 * zone) and set a Graph `transactionId` so retries do not create duplicates.
 */

import { createHash } from "node:crypto";
import { z } from "zod";
import { GraphError, ValidationError } from "../util/errors.js";
import { htmlToText } from "../util/html.js";
import {
  assertValidTimeZone,
  type GraphDateTime,
  normalizeLocalDateTime,
  preferTimeZoneHeader,
  toGraphDateTime,
  utcToZonedWallClock,
  zonedWallClockToUtcIso,
} from "../util/time.js";
import type { GraphClient, ODataPage } from "./client.js";
import { type MessageBody, sliceBody } from "./mail.js";
import { CursorBase, decodeCursor, encodeCursor } from "./paging.js";

const EVENT_SELECT =
  "id,subject,start,end,location,isAllDay,webLink,type,seriesMasterId,showAs,organizer,isCancelled";

/**
 * Where a calendar's events live: the default calendar, or the one
 * `calendarId` names (from list_calendars).
 */
function calendarPath(calendarId: string | undefined, rest: string): string {
  return calendarId ? `/me/calendars/${encodeURIComponent(calendarId)}/${rest}` : `/me/${rest}`;
}

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
  /** A calendar from list_calendars; the default calendar if omitted. */
  calendarId?: string;
  /** A previous call's nextPageToken; the other inputs are then ignored. */
  pageToken?: string;
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
  /** singleInstance | occurrence | exception | seriesMaster */
  type: string | null;
  /** The series an occurrence or exception belongs to. */
  seriesMasterId: string | null;
  /** free | tentative | busy | oof | workingElsewhere | unknown */
  showAs: string | null;
  organizer: { name?: string; address?: string } | null;
  isCancelled: boolean;
}

export interface ListEventsResult {
  events: EventSummary[];
  /** The zone the times are in (from the cursor when paging). */
  timeZone: string;
  /** True when the range has more events than were returned. */
  truncated: boolean;
  /** Pass as `pageToken` for the next events; present when there is a next page. */
  nextPageToken?: string;
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
  seriesMasterId?: string;
  showAs?: string;
  organizer?: { emailAddress?: { name?: string; address?: string } };
  isCancelled?: boolean;
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
    type: e.type ?? null,
    seriesMasterId: e.seriesMasterId ?? null,
    showAs: e.showAs ?? null,
    organizer: e.organizer?.emailAddress
      ? { name: e.organizer.emailAddress.name, address: e.organizer.emailAddress.address }
      : null,
    isCancelled: Boolean(e.isCancelled),
  };
}

// A list_events cursor (see ./paging.ts) also holds the display zone, since the
// `Prefer` header that renders it isn't part of Graph's nextLink. `kind` keeps
// a list_messages cursor from being followed here.
const EventsCursorSchema = CursorBase.extend({
  kind: z.literal("events"),
  timeZone: z.string().min(1),
});

/**
 * List events in a date range via calendarView, rendered in `timeZone`.
 *
 * `startDateTime`/`endDateTime` are ISO 8601. A bare local wall-clock value is
 * interpreted in `timeZone` (so the queried window matches the zone the results
 * are rendered in — a naive "today in Brisbane" query won't silently miss or
 * include boundary events); a value carrying an offset/`Z` pins the instant
 * directly. The `Prefer` header controls the zone of the times in the response.
 *
 * Returns one page of at most `limit` events. When the range holds more,
 * `truncated` is set (#49), so a capped list is never mistaken for the whole
 * range, and `nextPageToken` continues it.
 */
export async function listEvents(
  graph: GraphClient,
  input: ListEventsInput,
): Promise<ListEventsResult> {
  if (input.pageToken) {
    // The cursor alone decides the request, like list_messages.
    const cursor = decodeCursor(input.pageToken, EventsCursorSchema, "list_events");
    assertValidTimeZone(cursor.timeZone);
    return eventsPage(graph, cursor.next, undefined, cursor.timeZone);
  }

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

  return eventsPage(
    graph,
    calendarPath(input.calendarId, "calendarView"),
    {
      startDateTime,
      endDateTime,
      $select: EVENT_SELECT,
      $orderby: "start/dateTime",
      $top: input.limit,
    },
    input.timeZone,
  );
}

async function eventsPage(
  graph: GraphClient,
  path: string,
  query: Record<string, string | number> | undefined,
  timeZone: string,
): Promise<ListEventsResult> {
  const page = (
    await graph.request<ODataPage<GraphEvent>>("GET", path, {
      query,
      headers: { prefer: preferTimeZoneHeader(timeZone) },
    })
  ).data;
  const nextLink = page["@odata.nextLink"];
  const nextPageToken = encodeCursor(nextLink, { kind: "events", timeZone });
  return {
    events: (page.value ?? []).map(shapeEvent),
    timeZone,
    // Set even if the nextLink can't be followed: the list is still incomplete.
    truncated: Boolean(nextLink),
    ...(nextPageToken ? { nextPageToken } : {}),
  };
}

export interface GetEventInput {
  id: string;
  /** IANA zone for the returned times. */
  timeZone: string;
  /** The event's calendar, when it isn't the default one. */
  calendarId?: string;
  /** Most body characters to return (default: as for get_message). */
  maxChars?: number;
  /** Where in the (converted) body to start, to read past a previous cut. */
  startChar?: number;
}

export interface EventAttendee {
  name?: string;
  address?: string;
  /** required | optional | resource */
  type: string | null;
  /** none | organizer | tentativelyAccepted | accepted | declined | notResponded */
  response: string | null;
}

export interface EventDetail extends EventSummary {
  /** The notes, as plain text, capped like an email body (see `sliceBody`). */
  body: MessageBody;
  attendees: EventAttendee[];
  isReminderOn: boolean | null;
  reminderMinutesBeforeStart: number | null;
  /** Graph's `patternedRecurrence` ({ pattern, range }); only a series master has one. */
  recurrence: Record<string, unknown> | null;
}

interface GraphAttendee {
  type?: string;
  status?: { response?: string };
  emailAddress?: { name?: string; address?: string };
}

/**
 * Read one event in full (#50): its notes (HTML converted to text and capped,
 * like get_message), attendees, reminder and recurrence, with times rendered in
 * `timeZone`. An occurrence has no recurrence of its own: its series master
 * (`seriesMasterId`) does.
 */
export async function getEvent(graph: GraphClient, input: GetEventInput): Promise<EventDetail> {
  assertValidTimeZone(input.timeZone);
  const e = await graph.get<GraphEvent & { recurrence?: Record<string, unknown> | null }>(
    calendarPath(input.calendarId, `events/${encodeURIComponent(input.id)}`),
    {
      query: {
        $select: `${EVENT_SELECT},body,attendees,isReminderOn,reminderMinutesBeforeStart,recurrence`,
      },
      headers: { prefer: preferTimeZoneHeader(input.timeZone) },
    },
  );
  const raw = e.body?.content ?? "";
  const text = e.body?.contentType?.toLowerCase() === "html" ? htmlToText(raw) : raw;
  return {
    ...shapeEvent(e),
    body: sliceBody(text, "text", input.maxChars, input.startChar),
    attendees: ((e.attendees ?? []) as GraphAttendee[]).map((a) => ({
      name: a.emailAddress?.name,
      address: a.emailAddress?.address,
      type: a.type ?? null,
      response: a.status?.response ?? null,
    })),
    isReminderOn: e.isReminderOn ?? null,
    reminderMinutesBeforeStart: e.reminderMinutesBeforeStart ?? null,
    recurrence: e.recurrence ?? null,
  };
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
  /** A calendar from list_calendars; the default calendar if omitted. */
  calendarId?: string;
  /**
   * An all-day event (#51): start and end are dates (or midnights) in one
   * zone, and the end is exclusive, i.e. the day after the last day.
   */
  isAllDay?: boolean;
  /** How the time shows to others (#51): e.g. "free" for a travel-time block. */
  showAs?: ShowAs;
  /** Makes the event a recurring series (#51); see {@link buildRecurrence}. */
  recurrence?: RecurrenceInput;
}

export const SHOW_AS = ["free", "tentative", "busy", "oof", "workingElsewhere"] as const;
export type ShowAs = (typeof SHOW_AS)[number];

export const RECURRENCE_TYPES = [
  "daily",
  "weekly",
  "absoluteMonthly",
  "relativeMonthly",
  "absoluteYearly",
  "relativeYearly",
] as const;
export const DAYS_OF_WEEK = [
  "sunday",
  "monday",
  "tuesday",
  "wednesday",
  "thursday",
  "friday",
  "saturday",
] as const;
export const WEEK_INDEXES = ["first", "second", "third", "fourth", "last"] as const;
/** Most occurrences a "numbered" series may have. */
export const MAX_OCCURRENCES = 999;

/** A recurrence as the tool takes it: Graph's `patternedRecurrence`, minus its zone. */
export interface RecurrenceInput {
  pattern: {
    type: (typeof RECURRENCE_TYPES)[number];
    /** Every `interval` days / weeks / months / years. */
    interval: number;
    daysOfWeek?: (typeof DAYS_OF_WEEK)[number][];
    dayOfMonth?: number;
    month?: number;
    index?: (typeof WEEK_INDEXES)[number];
    firstDayOfWeek?: (typeof DAYS_OF_WEEK)[number];
  };
  range: {
    type: "endDate" | "noEnd" | "numbered";
    /** The first occurrence's date, YYYY-MM-DD: must be the event's start date. */
    startDate: string;
    endDate?: string;
    numberOfOccurrences?: number;
  };
}

/** A real calendar date `YYYY-MM-DD`, or a ValidationError naming `label`. */
function calendarDate(s: string, label: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    throw new ValidationError(`${label} must be a date "YYYY-MM-DD"; got "${s}".`);
  }
  normalizeLocalDateTime(`${s}T00:00:00`); // rejects 2026-02-30 and the like
  return s;
}

/**
 * Graph's `patternedRecurrence` for `r`, checked against the event's start.
 * Each pattern type gets exactly the fields it uses, and is refused without
 * them, so the series is the one described rather than whatever Graph makes of
 * a partial pattern. The range starts on the event's start date, in its zone.
 */
export function buildRecurrence(r: RecurrenceInput, start: GraphDateTime): Record<string, unknown> {
  const { pattern: p, range } = r;
  if (!Number.isInteger(p.interval) || p.interval < 1 || p.interval > 99) {
    throw new ValidationError("recurrence.pattern.interval must be a whole number from 1 to 99.");
  }
  const need = <T>(value: T | undefined, field: string): T => {
    if (value === undefined || (Array.isArray(value) && value.length === 0)) {
      throw new ValidationError(`A ${p.type} recurrence needs ${field}.`);
    }
    return value;
  };
  const pattern: Record<string, unknown> = { type: p.type, interval: p.interval };
  switch (p.type) {
    case "daily":
      break;
    case "weekly":
      pattern.daysOfWeek = need(p.daysOfWeek, "pattern.daysOfWeek");
      if (p.firstDayOfWeek) pattern.firstDayOfWeek = p.firstDayOfWeek;
      break;
    case "absoluteMonthly":
      pattern.dayOfMonth = need(p.dayOfMonth, "pattern.dayOfMonth");
      break;
    case "relativeMonthly":
      pattern.daysOfWeek = need(p.daysOfWeek, "pattern.daysOfWeek");
      pattern.index = need(p.index, "pattern.index");
      break;
    case "absoluteYearly":
      pattern.dayOfMonth = need(p.dayOfMonth, "pattern.dayOfMonth");
      pattern.month = need(p.month, "pattern.month");
      break;
    case "relativeYearly":
      pattern.daysOfWeek = need(p.daysOfWeek, "pattern.daysOfWeek");
      pattern.index = need(p.index, "pattern.index");
      pattern.month = need(p.month, "pattern.month");
      break;
  }

  const startDate = calendarDate(range.startDate, "recurrence.range.startDate");
  if (startDate !== start.dateTime.slice(0, 10)) {
    throw new ValidationError(
      `recurrence.range.startDate (${startDate}) must be the event's start date ` +
        `(${start.dateTime.slice(0, 10)}). Nothing was created.`,
    );
  }
  const out: Record<string, unknown> = {
    type: range.type,
    startDate,
    recurrenceTimeZone: start.timeZone,
  };
  if (range.type === "endDate") {
    if (range.endDate === undefined) {
      throw new ValidationError("An endDate recurrence range needs range.endDate.");
    }
    const endDate = calendarDate(range.endDate, "recurrence.range.endDate");
    if (endDate < startDate) {
      throw new ValidationError("recurrence.range.endDate must not be before its startDate.");
    }
    out.endDate = endDate;
  } else if (range.type === "numbered") {
    const n = range.numberOfOccurrences;
    if (n === undefined || !Number.isInteger(n) || n < 1 || n > MAX_OCCURRENCES) {
      throw new ValidationError(
        `A numbered recurrence needs range.numberOfOccurrences from 1 to ${MAX_OCCURRENCES}.`,
      );
    }
    out.numberOfOccurrences = n;
  }
  return { pattern, range: out };
}

/**
 * The start or end of an all-day event: midnight in its zone. A date
 * ("2026-07-20") means that midnight; a time must be midnight already.
 */
function allDayTime(t: EventDateTime, side: "start" | "end"): GraphDateTime {
  const local = /^\d{4}-\d{2}-\d{2}$/.test(t.dateTime.trim())
    ? `${t.dateTime.trim()}T00:00:00`
    : t.dateTime;
  const g = toGraphDateTime(local, t.timeZone);
  if (!g.dateTime.endsWith("T00:00:00")) {
    throw new ValidationError(
      `An all-day event's ${side} must be a date ("YYYY-MM-DD") or midnight; got ` +
        `"${t.dateTime}". Nothing was created.`,
    );
  }
  return g;
}

export interface UpdateEventInput {
  id: string;
  /** The event's calendar, when it isn't the default one. */
  calendarId?: string;
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
  "isAllDay",
  "showAs",
  "recurrence",
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
    // Only when set, so a plain event's id is what it always was.
    ...(input.isAllDay ? ["isAllDay"] : []),
    ...(input.recurrence ? [JSON.stringify(input.recurrence)] : []),
  ].join(String.fromCharCode(0));
  return createHash("sha256").update(material).digest("hex").slice(0, 36);
}

/**
 * Create a calendar event.
 *
 * Times are sent as local wall-clock + IANA time zone via {@link toGraphDateTime}
 * (which rejects offset/`Z` values), so the event lands at the intended local
 * time. When no `transactionId` is supplied a deterministic one is derived from
 * the event details, so a retried identical create does not duplicate. An
 * all-day event is midnight to midnight in one zone, and a recurrence is built
 * by {@link buildRecurrence} (#51).
 * `deriveId` is injectable for tests.
 */
export async function createEvent(
  graph: GraphClient,
  input: CreateEventInput,
  deriveId: (input: CreateEventInput) => string = deterministicTransactionId,
): Promise<CreatedEvent> {
  const transactionId = input.transactionId ?? deriveId(input);
  let start: GraphDateTime;
  let end: GraphDateTime;
  if (input.isAllDay) {
    // Graph requires midnight to midnight, in one zone (#51).
    start = allDayTime(input.start, "start");
    end = allDayTime(input.end, "end");
    if (start.timeZone !== end.timeZone) {
      throw new ValidationError(
        "An all-day event's start and end must be in the same time zone. Nothing was created.",
      );
    }
  } else {
    start = toGraphDateTime(input.start.dateTime, input.start.timeZone);
    end = toGraphDateTime(input.end.dateTime, input.end.timeZone);
  }
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
    ...(input.isAllDay ? { isAllDay: true } : {}),
    ...(input.showAs ? { showAs: input.showAs } : {}),
    ...(input.recurrence ? { recurrence: buildRecurrence(input.recurrence, start) } : {}),
  };
  assertWritableEventPayload(payload);
  const created = await graph.post<GraphEvent & { recurrence?: Record<string, unknown> | null }>(
    calendarPath(input.calendarId, "events"),
    payload,
  );
  return {
    ...shapeWrite(created, transactionId),
    isAllDay: Boolean(created.isAllDay),
    showAs: created.showAs ?? null,
    recurrence: created.recurrence ?? null,
  };
}

/** What create_event echoes: the written event, with the #51 fields. */
export interface CreatedEvent extends EventWriteResult {
  isAllDay: boolean;
  showAs: string | null;
  recurrence: Record<string, unknown> | null;
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
  const path = calendarPath(input.calendarId, `events/${encodeURIComponent(input.id)}`);
  // Changing either side needs both, to check the result still ends after it starts.
  const select = new Set(["id", "attendees", ...fields, ...(zone ? ["start", "end"] : [])]);

  // When times change, read the event in UTC: a stored time is then exact even
  // in a repeated daylight-saving hour, where the same wall-clock time happens
  // twice. The report still shows it in `zone`, like the updated event.
  const before = await graph.get<GraphEvent>(path, {
    query: { $select: [...select].join(",") },
    headers: zone ? { prefer: preferTimeZoneHeader("UTC") } : undefined,
  });
  if ((before.attendees?.length ?? 0) > 0) {
    throw new ValidationError(
      "This event has attendees, so changing it could email them an update. The connector " +
        "doesn't send mail; change it in Outlook instead. Nothing was changed.",
    );
  }
  if (zone) {
    // A side that isn't changing keeps its stored value (read in UTC above).
    const stored = (t?: { dateTime?: string }) =>
      t?.dateTime ? instant({ dateTime: t.dateTime, timeZone: "UTC" }) : undefined;
    const storedStart = stored(before.start);
    const storedEnd = stored(before.end);
    const startMs = input.start ? instant(patch.start as GraphDateTime) : storedStart;
    const endMs = input.end ? instant(patch.end as GraphDateTime) : storedEnd;
    if (startMs !== undefined && endMs !== undefined) {
      assertEndAfterStart(
        startMs,
        endMs,
        "After this change the event would end at or before it starts. Nothing was changed.",
      );
    }
    // Report the stored times in the zone of the change, as Graph renders the result.
    const inZone = (ms: number) => ({ dateTime: utcToZonedWallClock(ms, zone), timeZone: zone });
    if (storedStart !== undefined) before.start = inZone(storedStart);
    if (storedEnd !== undefined) before.end = inZone(storedEnd);
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
  /** The event's calendar, when it isn't the default one. */
  calendarId?: string;
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
  const path = calendarPath(input.calendarId, `events/${encodeURIComponent(input.id)}`);
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

// ── Calendars ────────────────────────────────────────────────────────────────

export interface CalendarSummary {
  id: string;
  name: string | null;
  isDefaultCalendar: boolean;
  /** Whether events can be created or changed in it. */
  canEdit: boolean;
  owner: { name?: string; address?: string } | null;
}

interface GraphCalendar {
  id: string;
  name?: string;
  isDefaultCalendar?: boolean;
  canEdit?: boolean;
  owner?: { name?: string; address?: string };
}

/** Most calendars list_calendars returns. */
export const MAX_CALENDARS = 100;

/** The user's calendars (`GET /me/calendars`), for list_events' `calendarId`. */
export async function listCalendars(graph: GraphClient): Promise<CalendarSummary[]> {
  const calendars = await graph.getAllPages<GraphCalendar>(
    "/me/calendars",
    { query: { $select: "id,name,isDefaultCalendar,canEdit,owner", $top: MAX_CALENDARS } },
    { maxItems: MAX_CALENDARS },
  );
  return calendars.map((c) => ({
    id: c.id,
    name: c.name ?? null,
    isDefaultCalendar: Boolean(c.isDefaultCalendar),
    canEdit: Boolean(c.canEdit),
    owner: c.owner ? { name: c.owner.name, address: c.owner.address } : null,
  }));
}
