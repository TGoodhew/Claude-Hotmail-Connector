/**
 * Calendar operations.
 *
 * Reads use `/me/calendarView` with a `Prefer: outlook.timezone` header so that
 * returned start/end times are rendered in the requested IANA zone (spec §7.3).
 * Write operations are added by a later issue.
 */

import { ValidationError } from "../util/errors.js";
import { assertValidTimeZone, preferTimeZoneHeader } from "../util/time.js";
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
