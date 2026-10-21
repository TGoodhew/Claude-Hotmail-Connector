/**
 * Timezone-correct helpers.
 *
 * The single most-called-out correctness risk in the spec is calendar writes
 * landing an hour off. Microsoft Graph expects a **local wall-clock** dateTime
 * string (no offset, no `Z`) paired with an **IANA** time zone, e.g.
 *   { dateTime: "2026-07-20T12:00:00", timeZone: "Australia/Brisbane" }
 * These helpers enforce exactly that shape and never convert between zones, so
 * the requested wall-clock time is preserved verbatim.
 */

import { ValidationError } from "./errors.js";

/** Graph `dateTimeTimeZone` resource shape. */
export interface GraphDateTime {
  /** Local wall-clock time, `YYYY-MM-DDTHH:mm:ss`, no offset/`Z`. */
  dateTime: string;
  /** IANA time zone name, e.g. `Australia/Brisbane`. */
  timeZone: string;
}

// `YYYY-MM-DDTHH:mm[:ss[.fraction]]`, with `T` or a space between date and time
// (models write both; Graph returns seven fractional digits).
const LOCAL_DT_RE = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?$/;
const DATE_ONLY_RE = /^(\d{4})-(\d{2})-(\d{2})$/;
const HAS_ZULU_RE = /[zZ]$/;
const HAS_OFFSET_RE = /[+-]\d{2}:?\d{2}$/;

/** True if `tz` is a time zone the runtime's ICU recognises (IANA names). */
export function isValidTimeZone(tz: string): boolean {
  if (!tz || typeof tz !== "string") return false;
  try {
    // Throws RangeError for unknown zones.
    new Intl.DateTimeFormat("en-US", { timeZone: tz });
    return true;
  } catch {
    return false;
  }
}

/** Throw {@link ValidationError} unless `tz` is a valid IANA zone. */
export function assertValidTimeZone(tz: string): void {
  if (!isValidTimeZone(tz)) {
    throw new ValidationError(
      `Unknown IANA time zone: "${tz}". Use a name like "Australia/Brisbane" or "UTC".`,
    );
  }
}

function assertRealCalendarDate(
  y: number,
  mo: number,
  d: number,
  h: number,
  mi: number,
  se: number,
  original: string,
): void {
  const dt = new Date(Date.UTC(y, mo - 1, d, h, mi, se));
  const ok =
    dt.getUTCFullYear() === y &&
    dt.getUTCMonth() === mo - 1 &&
    dt.getUTCDate() === d &&
    dt.getUTCHours() === h &&
    dt.getUTCMinutes() === mi &&
    dt.getUTCSeconds() === se;
  if (!ok) {
    throw new ValidationError(`"${original}" is not a real calendar date/time.`);
  }
}

/**
 * Validate and normalise a local wall-clock dateTime to `YYYY-MM-DDTHH:mm:ss`.
 *
 * Rejects any value carrying a UTC `Z` or a numeric offset — those conflict
 * with a named time zone and are the classic source of "event is an hour off".
 */
export function normalizeLocalDateTime(input: string): string {
  if (typeof input !== "string") {
    throw new ValidationError("dateTime must be a string.");
  }
  const s = input.trim();
  if (HAS_ZULU_RE.test(s) || HAS_OFFSET_RE.test(s)) {
    throw new ValidationError(
      `dateTime "${input}" must be a local wall-clock time WITHOUT a UTC "Z" or numeric offset; ` +
        `pair it with an IANA timeZone instead (e.g. "2026-07-20T12:00:00" + "Australia/Brisbane").`,
    );
  }
  // Event times are whole seconds, so a fraction is dropped here; callers check
  // ordering on this normalised value, which is what gets written.
  const { y, mo, d, h, mi, se } = parseLocal(s, input, false);
  const p2 = (n: number): string => String(n).padStart(2, "0");
  return `${String(y).padStart(4, "0")}-${p2(mo)}-${p2(d)}T${p2(h)}:${p2(mi)}:${p2(se)}`;
}

interface LocalParts {
  y: number;
  mo: number;
  d: number;
  h: number;
  mi: number;
  se: number;
  /** Milliseconds from the fraction (finer digits are dropped: JS dates hold ms). */
  ms: number;
}

/**
 * Parse and range-check a bare local date/time.
 * With `allowDateOnly`, `YYYY-MM-DD` means midnight that day. Anything else is
 * a {@link ValidationError}, never passed on for Graph to read as UTC (#33).
 */
function parseLocal(s: string, input: string, allowDateOnly: boolean): LocalParts {
  const m = LOCAL_DT_RE.exec(s) ?? (allowDateOnly ? DATE_ONLY_RE.exec(s) : null);
  if (!m) {
    throw new ValidationError(
      `Invalid local dateTime "${input}"; expected "YYYY-MM-DDTHH:mm" or "YYYY-MM-DDTHH:mm:ss"` +
        (allowDateOnly ? ', or a date "YYYY-MM-DD" for midnight.' : "."),
    );
  }
  const [y, mo, d, h, mi, se] = [1, 2, 3, 4, 5, 6].map((i) => Number(m[i] ?? 0)) as [
    number,
    number,
    number,
    number,
    number,
    number,
  ];
  if (mo < 1 || mo > 12) throw new ValidationError(`Month out of range in "${input}".`);
  if (d < 1 || d > 31) throw new ValidationError(`Day out of range in "${input}".`);
  if (h > 23) throw new ValidationError(`Hour out of range in "${input}".`);
  if (mi > 59) throw new ValidationError(`Minute out of range in "${input}".`);
  if (se > 59) throw new ValidationError(`Second out of range in "${input}".`);
  assertRealCalendarDate(y, mo, d, h, mi, se, input);
  const ms = Number((m[7] ?? "").padEnd(3, "0").slice(0, 3));
  return { y, mo, d, h, mi, se, ms };
}

/**
 * Build a Graph {@link GraphDateTime} from a local wall-clock string and an
 * IANA time zone. The wall-clock time is preserved verbatim (no conversion).
 */
export function toGraphDateTime(localDateTime: string, timeZone: string): GraphDateTime {
  assertValidTimeZone(timeZone);
  return { dateTime: normalizeLocalDateTime(localDateTime), timeZone };
}

/** True if an ISO string carries an explicit UTC `Z` or a numeric offset. */
export function hasExplicitOffset(s: string): boolean {
  const t = s.trim();
  return HAS_ZULU_RE.test(t) || HAS_OFFSET_RE.test(t);
}

/** Offset (ms) of `timeZone` at a given UTC instant: local wall-clock − UTC. */
function tzOffsetMs(utcMs: number, timeZone: string): number {
  const dtf = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hour12: false,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  const parts = dtf.formatToParts(new Date(utcMs));
  const get = (t: string): number => Number(parts.find((p) => p.type === t)?.value ?? "0");
  const hour = get("hour") % 24; // some ICU builds render midnight as "24"
  const asIfUtc = Date.UTC(
    get("year"),
    get("month") - 1,
    get("day"),
    hour,
    get("minute"),
    get("second"),
  );
  return asIfUtc - utcMs;
}

/**
 * Interpret a bare local wall-clock dateTime as a time in `timeZone` and return
 * the corresponding UTC instant as an ISO string (`...Z`).
 *
 * Used to pin a calendar query window so the queried range matches the zone the
 * results are rendered in, and to compare start and end. A value that already
 * carries an offset/`Z` is returned unchanged (once it parses). A date-only
 * value means local midnight. Anything else is a {@link ValidationError}: passed
 * through, Graph would read it as UTC and shift the window (#33).
 */
export function zonedWallClockToUtcIso(dateTime: string, timeZone: string): string {
  const s = dateTime.trim();
  if (hasExplicitOffset(s)) {
    if (Number.isNaN(Date.parse(s))) {
      throw new ValidationError(`Invalid date/time "${dateTime}".`);
    }
    return s;
  }
  assertValidTimeZone(timeZone);
  const p = parseLocal(s, dateTime, true);
  const asUtc = Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi, p.se);
  // The offset is found at whole seconds (zone formatting has no ms), then the
  // fraction is added back, so sub-second bounds keep their order.
  return new Date(localToUtcMs(asUtc, timeZone) + p.ms).toISOString();
}

const DAY_MS = 86_400_000;

/**
 * The UTC instant of a local wall-clock time (`wallMs`, the wall clock read as
 * if it were UTC) in `timeZone`. Around a DST change the zone's offset differs
 * before and after, so both are tried and only an instant that really shows
 * that wall-clock time is accepted. When the clocks go back and the time occurs
 * twice, the earlier instant is used. When they go forward and it never occurs
 * (e.g. midnight in a zone that skips 00:00-01:00), it maps to the moment the
 * clocks jump to, so a date-only bound is still the start of that local day
 * (the "compatible" rule, as in Temporal).
 */
function localToUtcMs(wallMs: number, timeZone: string): number {
  const offsets = [tzOffsetMs(wallMs - DAY_MS, timeZone), tzOffsetMs(wallMs + DAY_MS, timeZone)];
  const shows = (utc: number) => utc + tzOffsetMs(utc, timeZone) === wallMs;
  const valid = offsets.map((o) => wallMs - o).filter(shows);
  if (valid.length > 0) return Math.min(...valid);
  // In the gap: the earlier (smaller) offset gives the instant after the jump.
  return wallMs - Math.min(...offsets);
}

/**
 * Value for the `Prefer` request header so Graph returns calendar times in the
 * requested zone: `outlook.timezone="<IANA tz>"`.
 */
export function preferTimeZoneHeader(timeZone: string): string {
  assertValidTimeZone(timeZone);
  return `outlook.timezone="${timeZone}"`;
}
