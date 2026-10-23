/**
 * MCP tool definitions for the calendar: list_calendars, list_events, get_event
 * and find_free_time (reads), create_event and update_event (writes), and the
 * guarded cancel_event (soft delete). Event times are local wall-clock + IANA zone.
 */

import { z } from "zod";
import {
  cancelEvent,
  createEvent,
  DAYS_OF_WEEK,
  MAX_OCCURRENCES,
  RECURRENCE_TYPES,
  SHOW_AS,
  WEEK_INDEXES,
  getEvent,
  listCalendars,
  listEvents,
  updateEvent,
} from "../graph/calendar.js";
import { findFreeTime, MAX_RANGE_DAYS } from "../graph/freebusy.js";
import { DEFAULT_MAX_BODY_CHARS } from "../graph/mail.js";
import {
  defineTool,
  toolResult,
  toRegistered,
  UNTRUSTED_CONTENT_DESCRIPTION,
  untrustedResult,
  type RegisteredTool,
  type ToolContext,
  type ToolDefinition,
} from "./types.js";

/** A local wall-clock date/time with an optional IANA zone (defaults per config). */
const eventTimeShape = z.object({
  dateTime: z
    .string()
    .describe('Local wall-clock time, e.g. "2026-07-20T10:00:00" (no offset or Z).'),
  timeZone: z
    .string()
    .optional()
    .describe("IANA time zone (defaults to the server's DEFAULT_TIMEZONE)."),
});

/** Which calendar: one from list_calendars, or the default calendar when omitted. */
const calendarIdShape = z
  .string()
  .min(1)
  .optional()
  .describe("A calendar id from list_calendars. Omit for the default calendar.");

function resolveTime(
  t: { dateTime: string; timeZone?: string },
  ctx: ToolContext,
): { dateTime: string; timeZone: string } {
  return { dateTime: t.dateTime, timeZone: t.timeZone ?? ctx.config.defaultTimeZone };
}

const listEventsShape = {
  startDateTime: z
    .string()
    .describe(
      'Range start: a local "YYYY-MM-DDTHH:mm[:ss]" (or "YYYY-MM-DD" for midnight), read in ' +
        "timeZone; or an ISO 8601 value with an offset/Z for an exact instant.",
    ),
  endDateTime: z.string().describe("Range end, in the same forms. Must be after the start."),
  timeZone: z
    .string()
    .optional()
    .describe("IANA time zone for the returned times (defaults to the server's DEFAULT_TIMEZONE)."),
  limit: z.number().int().min(1).max(100).default(25).describe("Max events to return (1-100)."),
  calendarId: calendarIdShape,
  pageToken: z
    .string()
    .optional()
    .describe(
      "Opaque token from a previous call's nextPageToken, for the next events. The range, time " +
        "zone and calendar then come from the token; pass the other inputs unchanged.",
    ),
};

export const listEventsTool: ToolDefinition<typeof listEventsShape> = defineTool({
  name: "list_events",
  title: "List calendar events",
  description:
    "List events in the user's personal Outlook.com / Hotmail calendar in a date range (e.g. to " +
    "check for conflicts before adding a booking). " +
    "Times are returned in the requested IANA time zone. Each event says whether it is busy, " +
    "tentative or free (showAs), cancelled, and part of a recurring series (type, " +
    "seriesMasterId). If truncated is true the range has MORE events than returned: call again " +
    "with nextPageToken before concluding a time is free. Read-only. " +
    UNTRUSTED_CONTENT_DESCRIPTION,
  inputShape: listEventsShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  scopesAnyOf: ["Calendars.Read", "Calendars.ReadWrite"],
  async handler(input, ctx) {
    const { events, timeZone, truncated, nextPageToken } = await listEvents(ctx.graph, {
      startDateTime: input.startDateTime,
      endDateTime: input.endDateTime,
      timeZone: input.timeZone ?? ctx.config.defaultTimeZone,
      limit: input.limit,
      calendarId: input.calendarId,
      pageToken: input.pageToken,
    });
    const found =
      events.length === 0
        ? "No events found"
        : `Found ${events.length} event(s) (times in ${timeZone})`;
    const more = !truncated
      ? "."
      : nextPageToken
        ? "; more in this range (use nextPageToken)."
        : "; more in this range that can't be listed here (narrow the range).";
    return untrustedResult(`${found}${more}`, {
      count: events.length,
      timeZone,
      truncated,
      ...(nextPageToken ? { nextPageToken } : {}),
      events,
    });
  },
});

const getEventShape = {
  id: z.string().min(1).describe("The event id (from list_events)."),
  timeZone: z
    .string()
    .optional()
    .describe("IANA time zone for the returned times (defaults to the server's DEFAULT_TIMEZONE)."),
  calendarId: calendarIdShape,
  maxChars: z
    .number()
    .int()
    .min(1_000)
    .max(100_000)
    .default(DEFAULT_MAX_BODY_CHARS)
    .describe("Most characters of the event's notes to return (1,000-100,000, default 20,000)."),
  startChar: z
    .number()
    .int()
    .min(0)
    .default(0)
    .describe("Where in the notes to start: a previous result's body.nextStartChar."),
};

export const getEventTool: ToolDefinition<typeof getEventShape> = defineTool({
  name: "get_event",
  title: "Read a calendar event",
  description:
    "Read one event in the user's personal Outlook.com / Hotmail calendar in full: its notes " +
    "(as plain text), attendees and their responses, organizer, reminder, free/busy (showAs) " +
    "and, for a whole recurring series, its recurrence pattern. Use it to check an event " +
    "before update_event or cancel_event. An occurrence of a series has no pattern of its own: " +
    "read its seriesMasterId for that. Long notes are cut at maxChars: body.truncated is then " +
    "true, and startChar=body.nextStartChar reads on. Read-only. " +
    UNTRUSTED_CONTENT_DESCRIPTION,
  inputShape: getEventShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  scopesAnyOf: ["Calendars.Read", "Calendars.ReadWrite"],
  async handler(input, ctx) {
    const timeZone = input.timeZone ?? ctx.config.defaultTimeZone;
    const event = await getEvent(ctx.graph, {
      id: input.id,
      timeZone,
      calendarId: input.calendarId,
      maxChars: input.maxChars,
      startChar: input.startChar,
    });
    // The summary stays free of the event's own text; that is all in the data.
    const { body } = event;
    const part = body.truncated
      ? ` The notes are long: this is characters ${body.startChar ?? 0}-${body.nextStartChar} of ` +
        `${body.totalChars}; call again with startChar=${body.nextStartChar} for more.`
      : "";
    return untrustedResult(`Read the event (times in ${timeZone}); its details follow.${part}`, {
      timeZone,
      event,
    });
  },
});

const createEventShape = {
  subject: z.string().min(1).describe("Event title."),
  start: eventTimeShape.describe("Event start (local wall-clock + IANA time zone)."),
  end: eventTimeShape.describe("Event end (local wall-clock + IANA time zone)."),
  location: z.string().optional().describe("Location display name."),
  body: z.string().optional().describe("Plain-text notes/body."),
  isReminderOn: z.boolean().optional().describe("Whether a reminder is set."),
  reminderMinutesBeforeStart: z
    .number()
    .int()
    .min(0)
    .optional()
    .describe("Reminder lead time in minutes."),
  transactionId: z
    .string()
    .optional()
    .describe("Idempotency key; if omitted one is generated so retries don't duplicate."),
  calendarId: calendarIdShape,
  isAllDay: z
    .boolean()
    .optional()
    .describe(
      'All-day event: start and end are dates ("YYYY-MM-DD"), and end is the day AFTER the last ' +
        'day (a one-day event on 20 July is start "2026-07-20", end "2026-07-21").',
    ),
  showAs: z
    .enum(SHOW_AS)
    .optional()
    .describe(
      "How the time shows to others: timed events default to busy, all-day events to free. " +
        'E.g. "free" for a travel-time block.',
    ),
  recurrence: z
    .object({
      pattern: z.object({
        type: z.enum(RECURRENCE_TYPES),
        interval: z.number().int().min(1).max(99).describe("Every N days/weeks/months/years."),
        daysOfWeek: z
          .array(z.enum(DAYS_OF_WEEK))
          .min(1)
          .optional()
          .describe("weekly, relativeMonthly, relativeYearly: which days."),
        dayOfMonth: z
          .number()
          .int()
          .min(1)
          .max(31)
          .optional()
          .describe("absoluteMonthly, absoluteYearly: which day of the month."),
        month: z.number().int().min(1).max(12).optional().describe("Yearly patterns: which month."),
        index: z
          .enum(WEEK_INDEXES)
          .optional()
          .describe('Relative patterns: which week, e.g. "first" Monday.'),
        firstDayOfWeek: z.enum(DAYS_OF_WEEK).optional(),
      }),
      range: z.object({
        type: z.enum(["endDate", "noEnd", "numbered"]),
        startDate: z.string().describe("YYYY-MM-DD: must be the event's start date."),
        endDate: z.string().optional().describe('YYYY-MM-DD, for type "endDate" (inclusive).'),
        numberOfOccurrences: z
          .number()
          .int()
          .min(1)
          .max(MAX_OCCURRENCES)
          .optional()
          .describe('For type "numbered".'),
      }),
    })
    .optional()
    .describe(
      "Makes this a recurring series (Graph's patternedRecurrence). Each pattern type needs its " +
        "fields: weekly needs daysOfWeek; absoluteMonthly dayOfMonth; relativeMonthly daysOfWeek " +
        "and index; absoluteYearly dayOfMonth and month; relativeYearly daysOfWeek, index and month.",
    ),
};

export const createEventTool: ToolDefinition<typeof createEventShape> = defineTool({
  name: "create_event",
  title: "Create a calendar event",
  description:
    "Create an event in the user's personal Outlook.com / Hotmail calendar (e.g. a booking, " +
    "plus separate travel-time blocks). " +
    "Times are LOCAL wall-clock paired with an IANA time zone (never a UTC 'Z'), so the " +
    "event lands at the intended local time. A transactionId makes repeated calls idempotent " +
    "(no duplicates). Supports all-day events (isAllDay), free/busy (showAs) and recurring " +
    "series (recurrence). A recurring series can't be removed with cancel_event, only in " +
    "Outlook, so confirm its pattern and end with extra care. IMPORTANT: confirm the details " +
    "(title, date, start/end, timezone, location, any recurrence) with the user BEFORE calling, " +
    "then echo back the created event.",
  inputShape: createEventShape,
  annotations: {
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
  scopesAnyOf: ["Calendars.ReadWrite"],
  async handler(input, ctx) {
    const result = await createEvent(ctx.graph, {
      subject: input.subject,
      start: resolveTime(input.start, ctx),
      end: resolveTime(input.end, ctx),
      location: input.location,
      body: input.body,
      isReminderOn: input.isReminderOn,
      reminderMinutesBeforeStart: input.reminderMinutesBeforeStart,
      transactionId: input.transactionId,
      calendarId: input.calendarId,
      isAllDay: input.isAllDay,
      showAs: input.showAs,
      recurrence: input.recurrence,
    });
    const title = result.subject ?? input.subject;
    return toolResult(
      result.matchedExisting
        ? `An event with these details already exists ("${title}"): nothing new was created. ` +
            "It is in `event`; to create a second copy on purpose, pass a new transactionId."
        : `Created event "${title}".`,
      { event: result },
    );
  },
});

const updateEventShape = {
  id: z.string().min(1).describe("The event id to update (from list_events)."),
  subject: z.string().min(1).optional().describe("New title."),
  start: eventTimeShape.optional().describe("New start (local wall-clock + IANA time zone)."),
  end: eventTimeShape.optional().describe("New end (local wall-clock + IANA time zone)."),
  location: z.string().optional().describe("New location display name."),
  body: z.string().optional().describe("New plain-text body."),
  isReminderOn: z.boolean().optional(),
  reminderMinutesBeforeStart: z.number().int().min(0).optional(),
  calendarId: calendarIdShape,
};

export const updateEventTool: ToolDefinition<typeof updateEventShape> = defineTool({
  name: "update_event",
  title: "Update a calendar event",
  description:
    "Update an existing event in the user's personal Outlook.com / Hotmail calendar: its time, " +
    "location, subject, or body. Only the fields you " +
    "provide are changed, and they are OVERWRITTEN in place (Outlook keeps no history). Times use " +
    "LOCAL wall-clock + IANA time zone. Refuses events with attendees (changing a meeting could " +
    "email them an update). IMPORTANT: confirm the change with the user BEFORE calling. The " +
    "result lists each changed field's before and after value; show it to the user.",
  inputShape: updateEventShape,
  annotations: {
    readOnlyHint: false,
    // A PATCH overwrites existing field values in place (not additive).
    destructiveHint: true,
    idempotentHint: true,
    openWorldHint: true,
  },
  scopesAnyOf: ["Calendars.ReadWrite"],
  async handler(input, ctx) {
    const { event, changes } = await updateEvent(ctx.graph, {
      id: input.id,
      calendarId: input.calendarId,
      subject: input.subject,
      start: input.start ? resolveTime(input.start, ctx) : undefined,
      end: input.end ? resolveTime(input.end, ctx) : undefined,
      location: input.location,
      body: input.body,
      isReminderOn: input.isReminderOn,
      reminderMinutesBeforeStart: input.reminderMinutesBeforeStart,
    });
    const fields = changes.map((c) => c.field).join(", ");
    return toolResult(
      `Updated event "${event.subject ?? input.id}" (changed: ${fields}). ` +
        "Previous and new values are in `changes`.",
      { event, changes },
    );
  },
});

const cancelEventShape = {
  id: z.string().min(1).describe("The event id to cancel (from list_events)."),
  confirmSubject: z
    .string()
    .min(1)
    .describe(
      "The event's exact title, as confirmed with the user. Must match or nothing happens.",
    ),
  timeZone: z
    .string()
    .optional()
    .describe("IANA time zone for the echoed times (defaults to the server's DEFAULT_TIMEZONE)."),
  calendarId: calendarIdShape,
};

// Approved by the owner with guards (#55): soft delete only, title must match,
// no meetings with attendees, nothing recurring.
export const cancelEventTool: ToolDefinition<typeof cancelEventShape> = defineTool({
  name: "cancel_event",
  title: "Cancel (delete) a calendar event",
  description:
    "DESTRUCTIVE: deletes an event in the user's personal Outlook.com / Hotmail calendar (it moves to " +
    "Outlook's Deleted Items). IMPORTANT: only call after the user has explicitly confirmed which " +
    "event to delete. confirmSubject must be the event's exact title. Refuses events with attendees " +
    "(deleting a meeting could email cancellations) and recurring events. Echo the removed " +
    "event back to the user.",
  inputShape: cancelEventShape,
  annotations: {
    readOnlyHint: false,
    destructiveHint: true,
    // Repeating the call removes nothing further (the event is already gone).
    idempotentHint: true,
    openWorldHint: true,
  },
  scopesAnyOf: ["Calendars.ReadWrite"],
  async handler(input, ctx) {
    const timeZone = input.timeZone ?? ctx.config.defaultTimeZone;
    const event = await cancelEvent(ctx.graph, {
      id: input.id,
      calendarId: input.calendarId,
      confirmSubject: input.confirmSubject,
      timeZone,
    });
    return toolResult(
      `Cancelled event "${event.subject ?? input.id}" (moved to Outlook's Deleted Items).`,
      { cancelled: event },
    );
  },
});

const listCalendarsShape = {};

export const listCalendarsTool: ToolDefinition<typeof listCalendarsShape> = defineTool({
  name: "list_calendars",
  title: "List calendars",
  description:
    "List the calendars in the user's personal Outlook.com / Hotmail account (their own and any " +
    "shared with them), with which is the default and which can be edited. Use an id as " +
    "calendarId in the other calendar tools; omit calendarId for the default calendar. " +
    "Read-only. " +
    UNTRUSTED_CONTENT_DESCRIPTION,
  inputShape: listCalendarsShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  scopesAnyOf: ["Calendars.Read", "Calendars.ReadWrite"],
  async handler(_input, ctx) {
    const calendars = await listCalendars(ctx.graph);
    // Names of calendars shared by others are their text, hence untrusted.
    return untrustedResult(`Found ${calendars.length} calendar(s).`, {
      count: calendars.length,
      calendars,
    });
  },
});

const findFreeTimeShape = {
  startDateTime: z
    .string()
    .describe(
      'Range start: a local "YYYY-MM-DDTHH:mm[:ss]" (or "YYYY-MM-DD" for midnight), read in ' +
        "timeZone; or an ISO 8601 value with an offset/Z for an exact instant.",
    ),
  endDateTime: z
    .string()
    .describe(`Range end, in the same forms; at most ${MAX_RANGE_DAYS} days after the start.`),
  timeZone: z
    .string()
    .optional()
    .describe(
      "IANA time zone for the range, the hours and the slots (defaults to DEFAULT_TIMEZONE).",
    ),
  minDurationMinutes: z
    .number()
    .int()
    .min(5)
    .max(1440)
    .default(30)
    .describe("Shortest free slot to return, in minutes (5-1440, default 30)."),
  withinHours: z
    .object({
      start: z.string().describe('Local start of the daily window, "HH:mm", e.g. "09:00".'),
      end: z.string().describe('Local end of the daily window, "HH:mm", after the start.'),
    })
    .optional()
    .describe("Only look within these hours each day (e.g. working hours)."),
  tentativeIsBusy: z
    .boolean()
    .default(true)
    .describe("Treat tentative events as busy (default true)."),
  calendarId: calendarIdShape,
};

export const findFreeTimeTool: ToolDefinition<typeof findFreeTimeShape> = defineTool({
  name: "find_free_time",
  title: "Find free time",
  description:
    "Find free time in the user's personal Outlook.com / Hotmail calendar over a range (e.g. a " +
    "slot for a booking or travel), optionally only within daily hours. Returns free slots of at " +
    "least minDurationMinutes and the merged busy times, in the requested IANA time zone, each " +
    "also as an exact UTC instant (startUtc/endUtc: rely on those in the hour that repeats when " +
    "clocks go back). Events " +
    "shown as free, and cancelled ones, don't count as busy; tentative ones do unless " +
    "tentativeIsBusy is false. Returns times only, no event details (use list_events for those). " +
    "Read-only.",
  inputShape: findFreeTimeShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  scopesAnyOf: ["Calendars.Read", "Calendars.ReadWrite"],
  async handler(input, ctx) {
    const result = await findFreeTime(ctx.graph, {
      startDateTime: input.startDateTime,
      endDateTime: input.endDateTime,
      timeZone: input.timeZone ?? ctx.config.defaultTimeZone,
      minDurationMinutes: input.minDurationMinutes,
      withinHours: input.withinHours,
      tentativeIsBusy: input.tentativeIsBusy,
      calendarId: input.calendarId,
    });
    const n = result.freeSlots.length;
    const summary =
      n === 0
        ? `No free slot of ${input.minDurationMinutes}+ minutes in that range.`
        : `Found ${n} free slot(s) of ${input.minDurationMinutes}+ minutes (times in ${result.timeZone}).`;
    return toolResult(summary, result);
  },
});

export const calendarTools: RegisteredTool[] = [
  toRegistered(listCalendarsTool),
  toRegistered(listEventsTool),
  toRegistered(findFreeTimeTool),
  toRegistered(getEventTool),
  toRegistered(createEventTool),
  toRegistered(updateEventTool),
  toRegistered(cancelEventTool),
];
