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
  const d = DATE_RE.exec(value);
  if (d && (line.params.get("VALUE") ?? "DATE").toUpperCase() === "DATE") {
    if (!realDate(d[1]!, d[2]!, d[3]!)) return null;
    return { dateTime: `${d[1]}-${d[2]}-${d[3]}`, timeZone: null, allDay: true };
  }
  const t = DATE_TIME_RE.exec(value);
  if (!t) return null;
  const [, y, mo, da, h, mi, s, z] = t;
  if (!realDate(y!, mo!, da!)) return null;
  if (Number(h) > 23 || Number(mi) > 59 || Number(s) > 60) return null;
  const tzid = line.params.get("TZID")?.trim();
  return {
    dateTime: `${y}-${mo}-${da}T${h}:${mi}:${s === "60" ? "59" : s}`,
    timeZone: z ? "UTC" : tzid || null,
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
      if (component === "VEVENT" && event) {
        if (events.length < MAX_ICS_EVENTS) events.push(event);
        else truncated = true;
        event = null;
      }
      continue;
    }
    const inside = stack[stack.length - 1];
    if (inside === "VCALENDAR" && stack.length === 1 && line.name === "METHOD") {
      method = value.trim().toUpperCase() || null;
    }
    if (!event || inside !== "VEVENT") continue;
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
        event.duration = /^[+-]?P[0-9WDTHMS]+$/.test(value.trim()) ? value.trim() : null;
        break;
      case "STATUS":
        event.status = value.trim().toUpperCase() || null;
        break;
      case "RRULE":
        event.recurrence = capped(value.trim()) || null;
        break;
    }
  }
  return { method, events, truncated };
}
