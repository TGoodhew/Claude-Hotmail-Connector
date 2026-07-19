/**
 * MCP tool definitions for the calendar.
 *
 * Read tools live here now; write tools (create_event / update_event) are added
 * by a later issue and appended to {@link calendarTools}.
 */

import { z } from "zod";
import { createEvent, listEvents, updateEvent } from "../graph/calendar.js";
import {
  defineTool,
  toolResult,
  toRegistered,
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

function resolveTime(
  t: { dateTime: string; timeZone?: string },
  ctx: ToolContext,
): { dateTime: string; timeZone: string } {
  return { dateTime: t.dateTime, timeZone: t.timeZone ?? ctx.config.defaultTimeZone };
}

const listEventsShape = {
  startDateTime: z
    .string()
    .describe("Range start, ISO 8601. Include a timezone offset/Z to pin it precisely."),
  endDateTime: z.string().describe("Range end, ISO 8601."),
  timeZone: z
    .string()
    .optional()
    .describe("IANA time zone for the returned times (defaults to the server's DEFAULT_TIMEZONE)."),
  limit: z.number().int().min(1).max(100).default(25).describe("Max events to return (1-100)."),
};

export const listEventsTool: ToolDefinition<typeof listEventsShape> = defineTool({
  name: "list_events",
  title: "List calendar events",
  description:
    "List calendar events in a date range (e.g. to check for conflicts before adding a booking). " +
    "Times are returned in the requested IANA time zone. Read-only.",
  inputShape: listEventsShape,
  annotations: { readOnlyHint: true, openWorldHint: true },
  async handler(input, ctx) {
    const timeZone = input.timeZone ?? ctx.config.defaultTimeZone;
    const events = await listEvents(ctx.graph, {
      startDateTime: input.startDateTime,
      endDateTime: input.endDateTime,
      timeZone,
      limit: input.limit,
    });
    const summary =
      events.length === 0
        ? "No events found in that range."
        : `Found ${events.length} event(s) (times in ${timeZone}).`;
    return toolResult(summary, { count: events.length, timeZone, events });
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
};

export const createEventTool: ToolDefinition<typeof createEventShape> = defineTool({
  name: "create_event",
  title: "Create a calendar event",
  description:
    "Create a calendar event (e.g. a booking, plus separate travel-time blocks). " +
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
    });
    return toolResult(`Created event "${result.subject ?? input.subject}".`, {
      event: result as unknown as Record<string, unknown>,
    });
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
};

export const updateEventTool: ToolDefinition<typeof updateEventShape> = defineTool({
  name: "update_event",
  title: "Update a calendar event",
  description:
    "Update an existing calendar event's time, location, subject, or body. Only the fields you " +
    "provide are changed. Times use LOCAL wall-clock + IANA time zone. Confirm changes with the " +
    "user before calling.",
  inputShape: updateEventShape,
  annotations: {
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
  async handler(input, ctx) {
    const result = await updateEvent(ctx.graph, {
      id: input.id,
      subject: input.subject,
      start: input.start ? resolveTime(input.start, ctx) : undefined,
      end: input.end ? resolveTime(input.end, ctx) : undefined,
      location: input.location,
      body: input.body,
      isReminderOn: input.isReminderOn,
      reminderMinutesBeforeStart: input.reminderMinutesBeforeStart,
    });
    return toolResult(`Updated event "${result.subject ?? input.id}".`, {
      event: result as unknown as Record<string, unknown>,
    });
  },
});

// NOTE: cancel_event (event deletion) is deliberately NOT implemented. Per the
// spec it is a gated/destructive action requiring explicit owner approval, and
// the kick-off prompt says to stop and ask before adding it.

export const calendarTools: RegisteredTool[] = [
  toRegistered(listEventsTool),
  toRegistered(createEventTool),
  toRegistered(updateEventTool),
];
