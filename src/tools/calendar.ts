/**
 * MCP tool definitions for the calendar.
 *
 * Read tools live here now; write tools (create_event / update_event) are added
 * by a later issue and appended to {@link calendarTools}.
 */

import { z } from "zod";
import { listEvents } from "../graph/calendar.js";
import {
  defineTool,
  toolResult,
  toRegistered,
  type RegisteredTool,
  type ToolDefinition,
} from "./types.js";

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

export const calendarTools: RegisteredTool[] = [toRegistered(listEventsTool)];
