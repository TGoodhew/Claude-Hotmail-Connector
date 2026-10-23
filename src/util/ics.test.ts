import { describe, expect, it } from "vitest";
import { MAX_ICS_EVENTS, MAX_ICS_TEXT_CHARS, parseIcs } from "./ics.js";

const ics = (...lines: string[]) => lines.join("\r\n");

describe("parseIcs (#54)", () => {
  it("reads an invitation: times in a zone, place, text, method", () => {
    const cal = parseIcs(
      ics(
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "METHOD:REQUEST",
        "BEGIN:VTIMEZONE",
        "TZID:Australia/Brisbane",
        "BEGIN:STANDARD",
        "DTSTART:19700101T000000",
        "TZOFFSETFROM:+1000",
        "TZOFFSETTO:+1000",
        "END:STANDARD",
        "END:VTIMEZONE",
        "BEGIN:VEVENT",
        "UID:abc@example.com",
        "SUMMARY:Dinner\\, Monsoon",
        "DTSTART;TZID=Australia/Brisbane:20260720T190000",
        "DTEND;TZID=Australia/Brisbane:20260720T210000",
        'LOCATION;ALTREP="https://maps.example/a:b":12 Main St\\; Level 2',
        "DESCRIPTION:Table for 2.\\nBooking ref: X1\\\\Y2",
        "STATUS:CONFIRMED",
        "BEGIN:VALARM",
        "DESCRIPTION:Reminder",
        "TRIGGER:-PT15M",
        "END:VALARM",
        "END:VEVENT",
        "END:VCALENDAR",
      ),
    );
    expect(cal).toEqual({
      method: "REQUEST",
      truncated: false,
      events: [
        {
          summary: "Dinner, Monsoon",
          start: { dateTime: "2026-07-20T19:00:00", timeZone: "Australia/Brisbane", allDay: false },
          end: { dateTime: "2026-07-20T21:00:00", timeZone: "Australia/Brisbane", allDay: false },
          duration: null,
          location: "12 Main St; Level 2",
          // The alarm's DESCRIPTION isn't the event's.
          description: "Table for 2.\nBooking ref: X1\\Y2",
          status: "CONFIRMED",
          recurrence: null,
        },
      ],
    });
  });

  it("reads UTC times, all-day dates, floating times, durations and repeats", () => {
    const cal = parseIcs(
      [
        "BEGIN:VCALENDAR",
        "BEGIN:VEVENT",
        "SUMMARY:Flight",
        "DTSTART:20260801T223000Z",
        "DURATION:PT2H15M",
        "END:VEVENT",
        "BEGIN:VEVENT",
        "SUMMARY:Holiday",
        "DTSTART;VALUE=DATE:20260810",
        "DTEND;VALUE=DATE:20260815",
        "RRULE:FREQ=YEARLY",
        "END:VEVENT",
        "BEGIN:VEVENT",
        "DTSTART:20260901T090000",
        "END:VEVENT",
        "END:VCALENDAR",
      ].join("\n"), // bare LF line ends are read too
    );
    expect(cal.method).toBeNull();
    expect(cal.events.map((e) => [e.summary, e.start, e.end, e.duration, e.recurrence])).toEqual([
      [
        "Flight",
        { dateTime: "2026-08-01T22:30:00", timeZone: "UTC", allDay: false },
        null,
        "PT2H15M",
        null,
      ],
      [
        "Holiday",
        { dateTime: "2026-08-10", timeZone: null, allDay: true },
        { dateTime: "2026-08-15", timeZone: null, allDay: true },
        null,
        "FREQ=YEARLY",
      ],
      [null, { dateTime: "2026-09-01T09:00:00", timeZone: null, allDay: false }, null, null, null],
    ]);
  });

  it("unfolds long lines", () => {
    const cal = parseIcs(
      ics(
        "BEGIN:VCALENDAR",
        "BEGIN:VEVENT",
        "SUMMARY:A very long",
        "  summary that was",
        "\t folded",
        "END:VEVENT",
        "END:VCALENDAR",
      ),
    );
    expect(cal.events[0]!.summary).toBe("A very long summary that was folded");
  });

  it("skips what it doesn't understand rather than guess", () => {
    const cal = parseIcs(
      ics(
        "not a content line",
        "BEGIN:VEVENT", // outside a VCALENDAR: not read
        "SUMMARY:Stray",
        "END:VEVENT",
        "BEGIN:VCALENDAR",
        "BEGIN:VEVENT",
        "SUMMARY:Odd times",
        "DTSTART:2026-07-20 19:00", // not iCalendar form
        "DTEND:20261340T250000", // impossible
        "DURATION:soon",
        "X-WEIRD;;=:x",
        ":no name",
        "END:VEVENT",
        "END:VCALENDAR",
      ),
    );
    expect(cal.events).toEqual([
      {
        summary: "Odd times",
        start: null,
        end: null,
        duration: null,
        location: null,
        description: null,
        status: null,
        recurrence: null,
      },
    ]);
    expect(parseIcs("").events).toEqual([]);
    expect(parseIcs("\u0000\u0001binary").events).toEqual([]);
  });

  it("refuses impossible dates, and accepts leap days", () => {
    const starts = (...values: string[]) =>
      parseIcs(
        ics(
          "BEGIN:VCALENDAR",
          ...values.flatMap((v) => ["BEGIN:VEVENT", v, "END:VEVENT"]),
          "END:VCALENDAR",
        ),
      ).events.map((e) => e.start?.dateTime ?? null);
    expect(
      starts(
        "DTSTART;VALUE=DATE:20261340",
        "DTSTART;VALUE=DATE:20260230",
        "DTSTART:20260230T090000",
        "DTSTART:20250229T090000Z",
        "DTSTART;VALUE=DATE:20280229",
        "DTSTART:20280229T090000",
      ),
    ).toEqual([null, null, null, null, "2028-02-29", "2028-02-29T09:00:00"]);
  });

  it("ignores an END that doesn't close the innermost component", () => {
    const cal = parseIcs(
      ics(
        "BEGIN:VCALENDAR",
        "BEGIN:VEVENT",
        "DESCRIPTION:The real one",
        "BEGIN:VALARM",
        "END:X-UNKNOWN", // stray: the alarm is still open
        "DESCRIPTION:From the alarm",
        "END:VALARM",
        "SUMMARY:Still the event",
        "END:VEVENT",
        "END:VCALENDAR",
      ),
    );
    expect(cal.events).toHaveLength(1);
    expect(cal.events[0]).toMatchObject({
      description: "The real one",
      summary: "Still the event",
    });
  });

  it("ignores a VEVENT nested in another, keeping the outer one whole", () => {
    const cal = parseIcs(
      ics(
        "BEGIN:VCALENDAR",
        "BEGIN:VEVENT",
        "SUMMARY:Outer",
        "BEGIN:VEVENT",
        "SUMMARY:Nested",
        "LOCATION:Nowhere",
        "END:VEVENT",
        "LOCATION:Outer place",
        "END:VEVENT",
        "END:VCALENDAR",
      ),
    );
    expect(cal.events).toHaveLength(1);
    expect(cal.events[0]).toMatchObject({ summary: "Outer", location: "Outer place" });
  });

  it("accepts only RFC 5545 durations", () => {
    const duration = (v: string) =>
      parseIcs(
        ics("BEGIN:VCALENDAR", "BEGIN:VEVENT", `DURATION:${v}`, "END:VEVENT", "END:VCALENDAR"),
      ).events[0]!.duration;
    for (const ok of [
      "P2W",
      "P1D",
      "P1DT2H",
      "PT1H30M",
      "PT45M",
      "PT30S",
      "-PT15M",
      "+P1DT1H2M3S",
    ]) {
      expect(duration(ok), ok).toBe(ok);
    }
    for (const bad of ["PT", "P1", "P1M", "P1D2H", "P1Y", "PT1S2M", "P1W2D", "P", "1H"]) {
      expect(duration(bad), bad).toBeNull();
    }
  });

  it("caps the events and the text it returns", () => {
    const one = [
      "BEGIN:VEVENT",
      `DESCRIPTION:${"x".repeat(MAX_ICS_TEXT_CHARS + 10)}`,
      "END:VEVENT",
    ];
    const cal = parseIcs(
      ics(
        "BEGIN:VCALENDAR",
        ...Array.from({ length: MAX_ICS_EVENTS + 1 }, () => one).flat(),
        "END:VCALENDAR",
      ),
    );
    expect(cal.events).toHaveLength(MAX_ICS_EVENTS);
    expect(cal.truncated).toBe(true);
    expect(cal.events[0]!.description).toHaveLength(MAX_ICS_TEXT_CHARS + 1); // with the ellipsis
  });
});
