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

function formatFromUtcParts(dt: Date): string {
  const p2 = (n: number): string => String(n).padStart(2, "0");
  return (
    `${dt.getUTCFullYear()}-${p2(dt.getUTCMonth() + 1)}-${p2(dt.getUTCDate())}` +
    `T${p2(dt.getUTCHours())}:${p2(dt.getUTCMinutes())}:${p2(dt.getUTCSeconds())}`
  );
}

/**
 * Shift a local wall-clock time by a number of minutes (may be negative),
 * handling day/month rollover. This is pure wall-clock arithmetic — it does
 * NOT convert between zones — so it is safe for building travel-time blocks
 * around a booking (e.g. 15 minutes before/after).
 */
export function shiftLocalDateTime(localDateTime: string, minutes: number): string {
  const norm = normalizeLocalDateTime(localDateTime);
  const m = LOCAL_DT_RE.exec(norm)!;
  const dt = new Date(
    Date.UTC(
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3]),
      Number(m[4]),
      Number(m[5]),
      Number(m[6]),
    ),
  );
  dt.setUTCMinutes(dt.getUTCMinutes() + minutes);
  return formatFromUtcParts(dt);
}

/**
 * Value for the `Prefer` request header so Graph returns calendar times in the
 * requested zone: `outlook.timezone="<IANA tz>"`.
 */
export function preferTimeZoneHeader(timeZone: string): string {
  assertValidTimeZone(timeZone);
  return `outlook.timezone="${timeZone}"`;
}
