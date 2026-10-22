/**
 * MCP tool definitions for the calendar: list_calendars and list_events (reads),
 * create_event and update_event (writes), and the guarded cancel_event (soft
 * delete). Event times are local wall-clock + IANA zone.
 */

import { z } from "zod";
import {
  cancelEvent,
  createEvent,
  listCalendars,
  listEvents,
  updateEvent,
} from "../graph/calendar.js";
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
};

export const createEventTool: ToolDefinition<typeof createEventShape> = defineTool({
  name: "create_event",
  title: "Create a calendar event",
  description:
    "Create an event in the user's personal Outlook.com / Hotmail calendar (e.g. a booking, " +
    "plus separate travel-time blocks). " +
    "Times are LOCAL wall-clock paired with an IANA time zone (never a UTC 'Z'), so the " +
    "event lands at the intended local time. A transactionId makes repeated calls idempotent " +
    "(no duplicates). IMPORTANT: confirm the details (title, date, start/end, timezone, location) " +
    "with the user BEFORE calling, then echo back the created event.",
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
    });
    return toolResult(`Created event "${result.subject ?? input.subject}".`, { event: result });
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

export const calendarTools: RegisteredTool[] = [
  toRegistered(listCalendarsTool),
  toRegistered(listEventsTool),
  toRegistered(createEventTool),
  toRegistered(updateEventTool),
  toRegistered(cancelEventTool),
];
