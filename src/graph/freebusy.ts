/**
 * Free/busy lookup (#52): the gaps in a calendar range, worked out from the
 * events themselves (calendarView). Graph's `getSchedule` isn't reliably
 * available to personal accounts, and calendarView is what list_events
 * already uses, so the answer matches what the user sees there.
 */

import { ValidationError } from "../util/errors.js";
import { assertValidTimeZone, utcToZonedWallClock, zonedWallClockToUtcIso } from "../util/time.js";
import { listEvents } from "./calendar.js";
import type { GraphClient } from "./client.js";

export interface FindFreeTimeInput {
  /** Range bounds, in the same forms list_events takes (bare local times are read in `timeZone`). */
  startDateTime: string;
  endDateTime: string;
  /** IANA zone for the bounds, the daily hours and the returned slots. */
  timeZone: string;
  /** Shortest gap worth returning, in minutes. */
  minDurationMinutes: number;
  /** Only within these local hours each day ("HH:mm", end after start), e.g. 09:00-17:00. */
  withinHours?: { start: string; end: string };
  /** Count tentative events as busy (default true). */
  tentativeIsBusy?: boolean;
  /** A calendar from list_calendars; the default calendar if omitted. */
  calendarId?: string;
}

/**
 * Each endpoint comes twice: as local wall-clock time in the requested zone
 * ("YYYY-MM-DDTHH:mm:ss", the form create_event takes) and as the exact UTC
 * instant. In the hour that repeats when clocks go back, the local time alone
 * can't say which occurrence it means; the UTC one does.
 */
interface Endpoints {
  start: string;
  end: string;
  startUtc: string;
  endUtc: string;
}

export interface TimeSlot extends Endpoints {
  minutes: number;
}

export interface BusyInterval extends Endpoints {
  /** busy | tentative | oof | workingElsewhere */
  showAs: string;
}

export interface FindFreeTimeResult {
  timeZone: string;
  freeSlots: TimeSlot[];
  /** The busy time found, merged (no event text: only when, and how busy). */
  busy: BusyInterval[];
}

/** Longest range a lookup covers, so a lookup stays a bounded number of pages. */
export const MAX_RANGE_DAYS = 62;
/** Pages of events read at most; a range with more is refused rather than half-answered. */
const MAX_PAGES = 10;
const PAGE_SIZE = 100;
const MINUTE_MS = 60_000;
/**
 * How far past each end of the range events are read. All-day events are
 * matched by Graph on their dates as if those were UTC, so one on a local day
 * can sit up to 14 hours away from it; a day covers every zone (#158).
 */
const ALL_DAY_MARGIN_MS = 24 * 60 * MINUTE_MS;

const HHMM = /^([01]\d|2[0-3]):([0-5]\d)$/;

/** Whether an event's `showAs` takes the time. Free and unknown don't. */
function takesTime(showAs: string | null, tentativeIsBusy: boolean): boolean {
  switch (showAs) {
    case "busy":
    case "oof":
    case "workingElsewhere":
      return true;
    case "tentative":
      return tentativeIsBusy;
    default:
      return false;
  }
}

interface Interval {
  start: number;
  end: number;
}

/** How strongly a status takes the time: a merged span reports the strongest. */
const STRENGTH: Record<string, number> = { tentative: 1, workingElsewhere: 2, busy: 3, oof: 4 };

/**
 * Sorted, non-overlapping union of busy `intervals` (touching ones join),
 * each showing the strongest status among the events it covers.
 */
export function mergeBusy<T extends Interval & { showAs: string }>(intervals: T[]): T[] {
  const sorted = [...intervals].sort((a, b) => a.start - b.start);
  const out: T[] = [];
  for (const i of sorted) {
    const last = out[out.length - 1];
    if (last && i.start <= last.end) {
      last.end = Math.max(last.end, i.end);
      if ((STRENGTH[i.showAs] ?? 0) > (STRENGTH[last.showAs] ?? 0)) last.showAs = i.showAs;
    } else {
      out.push({ ...i });
    }
  }
  return out;
}

/** The parts of `windows` not covered by `busy` (both sorted and non-overlapping). */
export function subtractIntervals(windows: Interval[], busy: Interval[]): Interval[] {
  const free: Interval[] = [];
  for (const w of windows) {
    let cursor = w.start;
    for (const b of busy) {
      if (b.end <= cursor || b.start >= w.end) continue;
      if (b.start > cursor) free.push({ start: cursor, end: b.start });
      cursor = Math.max(cursor, b.end);
      if (cursor >= w.end) break;
    }
    if (cursor < w.end) free.push({ start: cursor, end: w.end });
  }
  return free;
}

/** The zone's UTC offset (ms) at an instant: local wall clock minus UTC. */
function offsetAt(ms: number, timeZone: string): number {
  return Date.parse(`${utcToZonedWallClock(ms, timeZone).slice(0, 23)}Z`) - ms;
}

/**
 * Split `range` into spans of constant UTC offset. The offset is sampled once a
 * day (zones change at most once a day), and each change found to the minute.
 */
function constantOffsetSpans(range: Interval, timeZone: string): (Interval & { offset: number })[] {
  const DAY_MS = 24 * 60 * MINUTE_MS;
  const spans: (Interval & { offset: number })[] = [];
  let start = range.start;
  let offset = offsetAt(start, timeZone);
  for (let t = range.start + DAY_MS; ; t += DAY_MS) {
    const probe = Math.min(t, range.end);
    if (offsetAt(probe, timeZone) !== offset) {
      // The first minute with the new offset, between start and probe.
      let lo = Math.floor(start / MINUTE_MS);
      let hi = Math.ceil(probe / MINUTE_MS);
      while (hi - lo > 1) {
        const mid = Math.floor((lo + hi) / 2);
        if (offsetAt(mid * MINUTE_MS, timeZone) === offset) lo = mid;
        else hi = mid;
      }
      const change = Math.max(start, Math.min(hi * MINUTE_MS, probe));
      spans.push({ start, end: change, offset });
      start = change;
      offset = offsetAt(change, timeZone);
    }
    if (probe >= range.end) break;
  }
  spans.push({ start, end: range.end, offset });
  return spans.filter((s) => s.end > s.start);
}

/**
 * The windows to look in: the whole range, or, with `withinHours`, exactly the
 * instants whose local clock reads within those hours. Worked out per span of
 * constant offset, so the hour skipped when clocks go forward contributes
 * nothing and the hour repeated when they go back counts twice, as two
 * separate windows; touching windows are then joined.
 */
function lookWindows(
  range: Interval,
  timeZone: string,
  withinHours: FindFreeTimeInput["withinHours"],
): Interval[] {
  if (!withinHours) return [range];
  const minutes = (hhmm: string) => Number(hhmm.slice(0, 2)) * 60 + Number(hhmm.slice(3, 5));
  const from = minutes(withinHours.start) * MINUTE_MS;
  const to = minutes(withinHours.end) * MINUTE_MS;
  const DAY_MS = 24 * 60 * MINUTE_MS;
  const pieces: Interval[] = [];
  for (const span of constantOffsetSpans(range, timeZone)) {
    // In this span, local wall clock (read as UTC) = instant + offset.
    const firstDay = Math.floor((span.start + span.offset) / DAY_MS) * DAY_MS;
    for (let day = firstDay; day - span.offset < span.end; day += DAY_MS) {
      const start = Math.max(day + from - span.offset, span.start);
      const end = Math.min(day + to - span.offset, span.end);
      if (end > start) pieces.push({ start, end });
    }
  }
  pieces.sort((a, b) => a.start - b.start);
  const windows: Interval[] = [];
  for (const p of pieces) {
    const last = windows[windows.length - 1];
    if (last && p.start <= last.end) last.end = Math.max(last.end, p.end);
    else windows.push({ ...p });
  }
  return windows;
}

/** Both forms of a result's endpoints (see {@link Endpoints}). */
function endpoints(start: number, end: number, timeZone: string): Endpoints {
  const local = (ms: number) => utcToZonedWallClock(ms, timeZone).slice(0, 19);
  const utc = (ms: number) => new Date(ms).toISOString().replace(".000Z", "Z");
  return { start: local(start), end: local(end), startUtc: utc(start), endUtc: utc(end) };
}

/**
 * Find free time in a range: read its events (in UTC, exactly), merge the time
 * they take, and return the gaps of at least `minDurationMinutes`, optionally
 * only within daily hours. Cancelled events and those shown as free take no
 * time; tentative ones do unless `tentativeIsBusy` is false.
 */
export async function findFreeTime(
  graph: GraphClient,
  input: FindFreeTimeInput,
): Promise<FindFreeTimeResult> {
  const { timeZone } = input;
  assertValidTimeZone(timeZone);
  const range = {
    start: Date.parse(zonedWallClockToUtcIso(input.startDateTime, timeZone)),
    end: Date.parse(zonedWallClockToUtcIso(input.endDateTime, timeZone)),
  };
  if (range.end <= range.start) {
    throw new ValidationError("endDateTime must be after startDateTime.");
  }
  if (range.end - range.start > MAX_RANGE_DAYS * 24 * 60 * MINUTE_MS) {
    throw new ValidationError(`The range can be at most ${MAX_RANGE_DAYS} days.`);
  }
  if (!Number.isInteger(input.minDurationMinutes) || input.minDurationMinutes < 1) {
    throw new ValidationError("minDurationMinutes must be a whole number of minutes, at least 1.");
  }
  if (input.withinHours) {
    const { start, end } = input.withinHours;
    if (!HHMM.test(start) || !HHMM.test(end) || end <= start) {
      throw new ValidationError(
        'withinHours needs "HH:mm" start and end times, with the end after the start.',
      );
    }
  }

  // Events in UTC, so their instants are exact whatever the zone does.
  const busy: (Interval & { showAs: string })[] = [];
  let pageToken: string | undefined;
  for (let page = 0; ; page++) {
    const result = await listEvents(graph, {
      startDateTime: new Date(range.start - ALL_DAY_MARGIN_MS).toISOString(),
      endDateTime: new Date(range.end + ALL_DAY_MARGIN_MS).toISOString(),
      timeZone: "UTC",
      limit: PAGE_SIZE,
      calendarId: input.calendarId,
      pageToken,
    });
    for (const e of result.events) {
      if (e.isCancelled || !takesTime(e.showAs, input.tentativeIsBusy ?? true)) continue;
      // Timed events come back in UTC. An all-day event's are dates, not
      // instants: Graph gives midnight to midnight in whatever zone was asked
      // for, so read them as local days in the requested zone (#158).
      const instant = (t: { dateTime: string | null } | null) => {
        if (!t?.dateTime) return NaN;
        if (!e.isAllDay) return Date.parse(`${t.dateTime}Z`);
        try {
          return Date.parse(zonedWallClockToUtcIso(t.dateTime.slice(0, 10), timeZone));
        } catch {
          return NaN;
        }
      };
      const start = instant(e.start);
      const end = instant(e.end);
      // An event that takes time but whose times can't be read may cover any
      // part of the range: don't call the range free on a guess.
      if (!Number.isFinite(start) || !Number.isFinite(end) || end < start) {
        throw new ValidationError(
          "An event in that range has times that couldn't be read, so free time can't be " +
            "worked out reliably. Check the calendar with list_events.",
        );
      }
      if (end > start) busy.push({ start, end, showAs: e.showAs ?? "busy" });
    }
    if (!result.truncated) break;
    if (!result.nextPageToken || page + 1 >= MAX_PAGES) {
      throw new ValidationError(
        "That range has too many events to work out free time reliably. Try a shorter range.",
      );
    }
    pageToken = result.nextPageToken;
  }

  const merged = mergeBusy(busy);
  const minMs = input.minDurationMinutes * MINUTE_MS;
  // Results are given to the second: shrink each slot to whole seconds inside
  // it, so an advertised start or end never overlaps busy time.
  const SECOND_MS = 1000;
  const free = subtractIntervals(lookWindows(range, timeZone, input.withinHours), merged)
    .map((s) => ({
      start: Math.ceil(s.start / SECOND_MS) * SECOND_MS,
      end: Math.floor(s.end / SECOND_MS) * SECOND_MS,
    }))
    .filter((s) => s.end - s.start >= minMs);
  return {
    timeZone,
    freeSlots: free.map((s) => ({
      ...endpoints(s.start, s.end, timeZone),
      // Whole minutes actually free: rounding up could promise time past the slot's end.
      minutes: Math.floor((s.end - s.start) / MINUTE_MS),
    })),
    busy: merged
      .filter((b) => b.end > range.start && b.start < range.end)
      .map((b) => ({
        ...endpoints(Math.max(b.start, range.start), Math.min(b.end, range.end), timeZone),
        showAs: b.showAs,
      })),
  };
}
