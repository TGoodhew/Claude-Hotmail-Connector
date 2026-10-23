/**
 * A small, strict iCalendar (RFC 5545) reader for `.ics` attachments (#54):
 * just the events, with their times, place and text. Anything it doesn't
 * understand is skipped rather than guessed at; it never evaluates recurrence
 * rules, it only reports them.
 */

/** A time as the file gives it: local wall-clock in a zone, a UTC instant, or a date. */
export interface IcsTime {
  /** "YYYY-MM-DD" for an all-day date, else "YYYY-MM-DDTHH:mm:ss". */
  dateTime: string;
  /**
   * The zone the time is in: "UTC", the file's TZID (often IANA, sometimes a
   * Windows name), or null for a date, or a "floating" time with no zone.
   */
  timeZone: string | null;
  allDay: boolean;
}

export interface IcsEvent {
  summary: string | null;
  start: IcsTime | null;
  end: IcsTime | null;
  /** The event's DURATION (e.g. "PT1H30M"), when it gives one instead of an end. */
  duration: string | null;
  location: string | null;
  description: string | null;
  /** TENTATIVE | CONFIRMED | CANCELLED, as the file says. */
  status: string | null;
  /** The RRULE as written, when the event repeats. Not expanded. */
  recurrence: string | null;
}

export interface IcsCalendar {
  /** The calendar's METHOD, e.g. REQUEST (an invitation) or CANCEL. */
  method: string | null;
  events: IcsEvent[];
  /** True when the file had more than {@link MAX_ICS_EVENTS} events. */
  truncated: boolean;
}

/** Most events returned from one file. */
export const MAX_ICS_EVENTS = 50;
/** Longest description kept per event; longer ones are cut. */
export const MAX_ICS_TEXT_CHARS = 4_000;
/**
 * Longest identifier-like value kept (METHOD, STATUS, TZID, DURATION). These are
 * never cut, since a cut one would mean something else: longer ones are dropped.
 */
export const MAX_ICS_TOKEN_CHARS = 100;
/** Longest RRULE kept, likewise never cut. */
export const MAX_ICS_RRULE_CHARS = 1_000;

/** The statuses a VEVENT may have (RFC 5545, 3.8.1.11). */
const EVENT_STATUSES = new Set(["TENTATIVE", "CONFIRMED", "CANCELLED"]);
/** An iana-token, as METHOD values are. */
const TOKEN_RE = /^[A-Za-z0-9-]+$/;

/** `value` trimmed, or null when it's empty or longer than `max`. */
function bounded(value: string, max = MAX_ICS_TOKEN_CHARS): string | null {
  const v = value.trim();
  return v && v.length <= max ? v : null;
}

interface ContentLine {
  name: string;
  params: Map<string, string>;
  value: string;
}

/** Undo line folding: a line starting with a space or tab continues the one before. */
function unfold(text: string): string[] {
  const lines: string[] = [];
  for (const raw of text.split(/\r\n|\n|\r/)) {
    if ((raw.startsWith(" ") || raw.startsWith("\t")) && lines.length > 0) {
      lines[lines.length - 1] += raw.slice(1);
    } else {
      lines.push(raw);
    }
  }
  return lines;
}

/** `NAME;PARAM=a;PARAM2="b:c":value`, or null if it isn't a content line. */
function parseLine(line: string): ContentLine | null {
  // The value starts at the first colon outside a quoted parameter value.
  let inQuotes = false;
  let colon = -1;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') inQuotes = !inQuotes;
    else if (c === ":" && !inQuotes) {
      colon = i;
      break;
    }
  }
  if (colon <= 0) return null;
  const head = line.slice(0, colon);
  const parts: string[] = [];
  let current = "";
  inQuotes = false;
  for (const c of head) {
    if (c === '"') inQuotes = !inQuotes;
    if (c === ";" && !inQuotes) {
      parts.push(current);
      current = "";
    } else {
      current += c;
    }
  }
  parts.push(current);
  const name = parts[0]!.trim().toUpperCase();
  if (!/^[A-Z0-9-]+$/.test(name)) return null;
  const params = new Map<string, string>();
  for (const p of parts.slice(1)) {
    const eq = p.indexOf("=");
    if (eq <= 0) continue;
    params.set(p.slice(0, eq).trim().toUpperCase(), p.slice(eq + 1).replace(/^"|"$/g, ""));
  }
  return { name, params, value: line.slice(colon + 1) };
}

/** A TEXT value, unescaped (`\n`, `\,`, `\;`, `\\`). */
function unescapeText(value: string): string {
  return value.replace(/\\([\\;,nN])/g, (_, c: string) => (c === "n" || c === "N" ? "\n" : c));
}

function capped(s: string): string {
  return s.length > MAX_ICS_TEXT_CHARS ? `${s.slice(0, MAX_ICS_TEXT_CHARS)}…` : s;
}

// RFC 5545 dur-value: weeks, or days and/or a time part (hours, minutes,
// seconds in order, each only after the one before it). No months or years.
const DUR_TIME = String.raw`T(?:\d+H(?:\d+M(?:\d+S)?)?|\d+M(?:\d+S)?|\d+S)`;
const DURATION_RE = new RegExp(String.raw`^[+-]?P(?:\d+W|\d+D(?:${DUR_TIME})?|${DUR_TIME})$`);
const DATE_RE = /^(\d{4})(\d{2})(\d{2})$/;
const DATE_TIME_RE = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/;

/** A DTSTART / DTEND value, or null if it isn't one this reader accepts. */
/** Whether the year, month and day make a real calendar date (leap years included). */
function realDate(y: string, mo: string, d: string): boolean {
  const date = new Date(Date.UTC(Number(y), Number(mo) - 1, Number(d)));
  return (
    date.getUTCFullYear() === Number(y) &&
    date.getUTCMonth() === Number(mo) - 1 &&
    date.getUTCDate() === Number(d)
  );
}

function parseTime(line: ContentLine): IcsTime | null {
  const value = line.value.trim();
  // A declared value type must be one this reader handles, and match the value.
  const type = line.params.get("VALUE")?.trim().toUpperCase();
  if (type !== undefined && type !== "DATE" && type !== "DATE-TIME") return null;
  // An oversized zone name isn't one: the value is then of no known zone,
  // whichever form it takes.
  const tzid = line.params.get("TZID");
  if (tzid !== undefined && !bounded(tzid)) return null;
  const d = DATE_RE.exec(value);
  // A bare date with no VALUE is strictly malformed (the default type is
  // DATE-TIME), but it's unambiguous and booking systems do send it, so it's
  // read as the date it plainly is.
  if (d && type !== "DATE-TIME") {
    if (!realDate(d[1]!, d[2]!, d[3]!)) return null;
    return { dateTime: `${d[1]}-${d[2]}-${d[3]}`, timeZone: null, allDay: true };
  }
  const t = DATE_TIME_RE.exec(value);
  if (!t || type === "DATE") return null;
  const [, y, mo, da, h, mi, s, z] = t;
  if (!realDate(y!, mo!, da!)) return null;
  if (Number(h) > 23 || Number(mi) > 59 || Number(s) > 60) return null;
  return {
    dateTime: `${y}-${mo}-${da}T${h}:${mi}:${s === "60" ? "59" : s}`,
    timeZone: z ? "UTC" : (tzid?.trim() ?? null),
    allDay: false,
  };
}

/** Read the events from an iCalendar file's text. */
export function parseIcs(text: string): IcsCalendar {
  const events: IcsEvent[] = [];
  let method: string | null = null;
  let truncated = false;
  // The components we're inside, innermost last; an event's own properties are
  // those met while VEVENT is innermost (not its VALARMs').
  const stack: string[] = [];
  let event: IcsEvent | null = null;

  for (const raw of unfold(text)) {
    const line = parseLine(raw);
    if (!line) continue;
    const value = line.value;
    if (line.name === "BEGIN") {
      const component = value.trim().toUpperCase();
      stack.push(component);
      if (component === "VEVENT" && stack.length === 2 && stack[0] === "VCALENDAR") {
        event = {
          summary: null,
          start: null,
          end: null,
          duration: null,
          location: null,
          description: null,
          status: null,
          recurrence: null,
        };
      }
      continue;
    }
    if (line.name === "END") {
      // Only the innermost open component can close; a stray END is ignored,
      // so nothing after it is credited to the wrong component.
      if (value.trim().toUpperCase() !== stack[stack.length - 1]) continue;
      const component = stack.pop();
      // Only a VEVENT directly inside the VCALENDAR is an event; one nested
      // deeper (malformed) neither fills nor closes it.
      if (component === "VEVENT" && event && stack.length === 1) {
        if (events.length < MAX_ICS_EVENTS) events.push(event);
        else truncated = true;
        event = null;
      }
      continue;
    }
    const inside = stack[stack.length - 1];
    if (inside === "VCALENDAR" && stack.length === 1 && line.name === "METHOD") {
      const m = bounded(value);
      method = m && TOKEN_RE.test(m) ? m.toUpperCase() : null;
    }
    if (!event || stack.length !== 2 || inside !== "VEVENT") continue;
    switch (line.name) {
      case "SUMMARY":
        event.summary = capped(unescapeText(value));
        break;
      case "LOCATION":
        event.location = capped(unescapeText(value));
        break;
      case "DESCRIPTION":
        event.description = capped(unescapeText(value));
        break;
      case "DTSTART":
        event.start = parseTime(line);
        break;
      case "DTEND":
        event.end = parseTime(line);
        break;
      case "DURATION":
        event.duration = ((d) => (d && DURATION_RE.test(d) ? d : null))(bounded(value));
        break;
      case "STATUS":
        event.status = ((st) => (st && EVENT_STATUSES.has(st) ? st : null))(
          bounded(value)?.toUpperCase(),
        );
        break;
      case "RRULE":
        event.recurrence = bounded(value, MAX_ICS_RRULE_CHARS);
        break;
    }
  }
  return { method, events, truncated };
}
