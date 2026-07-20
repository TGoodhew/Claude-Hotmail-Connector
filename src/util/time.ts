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

const LOCAL_DT_RE = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/;
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
  const m = LOCAL_DT_RE.exec(s);
  if (!m) {
    throw new ValidationError(
      `Invalid local dateTime "${input}"; expected "YYYY-MM-DDTHH:mm" or "YYYY-MM-DDTHH:mm:ss".`,
    );
  }
  const [, y, mo, d, h, mi, se] = m;
  const yy = Number(y);
  const moo = Number(mo);
  const dd = Number(d);
  const hh = Number(h);
  const mii = Number(mi);
  const sss = se ? Number(se) : 0;
  if (moo < 1 || moo > 12) throw new ValidationError(`Month out of range in "${input}".`);
  if (dd < 1 || dd > 31) throw new ValidationError(`Day out of range in "${input}".`);
  if (hh > 23) throw new ValidationError(`Hour out of range in "${input}".`);
  if (mii > 59) throw new ValidationError(`Minute out of range in "${input}".`);
  if (sss > 59) throw new ValidationError(`Second out of range in "${input}".`);
  assertRealCalendarDate(yy, moo, dd, hh, mii, sss, input);
  const p2 = (n: number): string => String(n).padStart(2, "0");
  return `${y}-${mo}-${d}T${p2(hh)}:${p2(mii)}:${p2(sss)}`;
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
  const asIfUtc = Date.UTC(get("year"), get("month") - 1, get("day"), hour, get("minute"), get("second"));
  return asIfUtc - utcMs;
}

/**
 * Interpret a bare local wall-clock dateTime as a time in `timeZone` and return
 * the corresponding UTC instant as an ISO string (`...Z`).
 *
 * Used to pin a calendar query window so the queried range matches the zone the
 * results are rendered in. If the input already carries an offset/`Z` it is
 * returned unchanged; anything that isn't a bare `YYYY-MM-DDTHH:mm[:ss]` (e.g. a
 * date-only value) is passed through untouched for Graph to interpret.
 */
export function zonedWallClockToUtcIso(dateTime: string, timeZone: string): string {
  const s = dateTime.trim();
  if (hasExplicitOffset(s)) return s;
  const m = LOCAL_DT_RE.exec(s);
  if (!m) return dateTime;
  assertValidTimeZone(timeZone);
  const asUtc = Date.UTC(
    Number(m[1]),
    Number(m[2]) - 1,
    Number(m[3]),
    Number(m[4]),
    Number(m[5]),
    m[6] ? Number(m[6]) : 0,
  );
  let offset = tzOffsetMs(asUtc, timeZone);
  // Refine once so DST boundaries (where the offset differs at the corrected
  // instant) resolve to the correct UTC time.
  const refined = tzOffsetMs(asUtc - offset, timeZone);
  if (refined !== offset) offset = refined;
  return new Date(asUtc - offset).toISOString();
}

/**
 * Value for the `Prefer` request header so Graph returns calendar times in the
 * requested zone: `outlook.timezone="<IANA tz>"`.
 */
export function preferTimeZoneHeader(timeZone: string): string {
  assertValidTimeZone(timeZone);
  return `outlook.timezone="${timeZone}"`;
}
