import { describe, expect, it } from "vitest";
import { ValidationError } from "./errors.js";
import {
  hasExplicitOffset,
  isValidTimeZone,
  normalizeLocalDateTime,
  preferTimeZoneHeader,
  toGraphDateTime,
  utcToZonedWallClock,
  zonedWallClockToUtcIso,
} from "./time.js";

describe("isValidTimeZone", () => {
  it("accepts real IANA zones", () => {
    expect(isValidTimeZone("Australia/Brisbane")).toBe(true);
    expect(isValidTimeZone("UTC")).toBe(true);
    expect(isValidTimeZone("Europe/London")).toBe(true);
  });
  it("rejects nonsense and empty input", () => {
    expect(isValidTimeZone("Mars/Phobos")).toBe(false);
    expect(isValidTimeZone("")).toBe(false);
    expect(isValidTimeZone("Australia/Nowhere")).toBe(false);
  });
});

describe("normalizeLocalDateTime", () => {
  it("preserves a full local wall-clock time", () => {
    expect(normalizeLocalDateTime("2026-07-20T12:00:00")).toBe("2026-07-20T12:00:00");
  });
  it("pads a missing seconds component", () => {
    expect(normalizeLocalDateTime("2026-07-20T12:00")).toBe("2026-07-20T12:00:00");
  });
  it("rejects a UTC 'Z' value (would conflict with a named zone)", () => {
    expect(() => normalizeLocalDateTime("2026-07-20T12:00:00Z")).toThrow(ValidationError);
  });
  it("rejects a numeric offset", () => {
    expect(() => normalizeLocalDateTime("2026-07-20T12:00:00+10:00")).toThrow(ValidationError);
    expect(() => normalizeLocalDateTime("2026-07-20T12:00:00-05:00")).toThrow(ValidationError);
  });
  it("rejects malformed and impossible dates", () => {
    expect(() => normalizeLocalDateTime("not-a-date")).toThrow(ValidationError);
    expect(() => normalizeLocalDateTime("2026-02-30T00:00:00")).toThrow(ValidationError);
    expect(() => normalizeLocalDateTime("2026-13-01T00:00:00")).toThrow(ValidationError);
    expect(() => normalizeLocalDateTime("2026-07-20T25:00:00")).toThrow(ValidationError);
  });
});

describe("toGraphDateTime — timezone correctness (spec §13)", () => {
  it("preserves 2026-07-20T12:00:00 + Australia/Brisbane without shifting", () => {
    const g = toGraphDateTime("2026-07-20T12:00:00", "Australia/Brisbane");
    expect(g).toEqual({ dateTime: "2026-07-20T12:00:00", timeZone: "Australia/Brisbane" });
    // Critically: the dateTime is NOT converted to UTC (would be 02:00:00Z).
    expect(g.dateTime).not.toContain("Z");
    expect(g.dateTime).not.toBe("2026-07-20T02:00:00");
  });
  it("rejects an unknown time zone", () => {
    expect(() => toGraphDateTime("2026-07-20T12:00:00", "Mars/Phobos")).toThrow(ValidationError);
  });
});

describe("hasExplicitOffset", () => {
  it("detects Z and numeric offsets", () => {
    expect(hasExplicitOffset("2026-07-20T00:00:00Z")).toBe(true);
    expect(hasExplicitOffset("2026-07-20T00:00:00+10:00")).toBe(true);
    expect(hasExplicitOffset("2026-07-20T00:00:00-0500")).toBe(true);
  });
  it("returns false for a bare local wall-clock", () => {
    expect(hasExplicitOffset("2026-07-20T00:00:00")).toBe(false);
    expect(hasExplicitOffset("2026-07-20T00:00")).toBe(false);
  });
});

describe("zonedWallClockToUtcIso — calendar window pinning", () => {
  it("interprets a bare local time in the given zone (Brisbane = UTC+10, no DST)", () => {
    // 2026-07-20 00:00 in Brisbane is 2026-07-19 14:00 UTC.
    expect(zonedWallClockToUtcIso("2026-07-20T00:00:00", "Australia/Brisbane")).toBe(
      "2026-07-19T14:00:00.000Z",
    );
  });
  it("is a no-op when the value already carries an offset/Z", () => {
    expect(zonedWallClockToUtcIso("2026-07-20T00:00:00Z", "Australia/Brisbane")).toBe(
      "2026-07-20T00:00:00Z",
    );
    expect(zonedWallClockToUtcIso("2026-07-20T00:00:00+10:00", "Australia/Brisbane")).toBe(
      "2026-07-20T00:00:00+10:00",
    );
  });
  it("handles a DST zone on both sides of the transition (New York)", () => {
    // July: EDT = UTC-4, so 12:00 local is 16:00 UTC.
    expect(zonedWallClockToUtcIso("2026-07-20T12:00:00", "America/New_York")).toBe(
      "2026-07-20T16:00:00.000Z",
    );
    // January: EST = UTC-5, so 12:00 local is 17:00 UTC.
    expect(zonedWallClockToUtcIso("2026-01-20T12:00:00", "America/New_York")).toBe(
      "2026-01-20T17:00:00.000Z",
    );
  });
  // #33: these shapes used to reach Graph unpinned, where they read as UTC and
  // shifted a Brisbane window by 10 hours.
  it("reads a date-only value as local midnight in the zone", () => {
    expect(zonedWallClockToUtcIso("2026-07-20", "Australia/Brisbane")).toBe(
      "2026-07-19T14:00:00.000Z",
    );
    // New York on the spring-forward day: midnight is still EST (UTC-5).
    expect(zonedWallClockToUtcIso("2026-03-08", "America/New_York")).toBe(
      "2026-03-08T05:00:00.000Z",
    );
  });
  it("accepts fractional seconds and a space separator", () => {
    for (const v of [
      "2026-07-20T10:00:00.000",
      "2026-07-20T10:00:00.0000000",
      "2026-07-20 10:00",
    ]) {
      expect(zonedWallClockToUtcIso(v, "Australia/Brisbane")).toBe("2026-07-20T00:00:00.000Z");
    }
  });
  it("keeps a nonzero fraction to the millisecond (Copilot review on #116)", () => {
    expect(zonedWallClockToUtcIso("2026-07-20T10:00:00.1", "Australia/Brisbane")).toBe(
      "2026-07-20T00:00:00.100Z",
    );
    expect(zonedWallClockToUtcIso("2026-07-20T10:00:00.9876543", "Australia/Brisbane")).toBe(
      "2026-07-20T00:00:00.987Z",
    );
    // Across a DST change the fraction still lands on the right side of the second.
    expect(zonedWallClockToUtcIso("2026-03-08T03:00:00.500", "America/New_York")).toBe(
      "2026-03-08T07:00:00.500Z",
    );
  });
  it("rejects anything else instead of passing it on", () => {
    for (const v of ["next tuesday", "07/20/2026", "2026-07-20T25:00", "2026-02-30", "2026-07"]) {
      expect(() => zonedWallClockToUtcIso(v, "Australia/Brisbane"), v).toThrow(ValidationError);
    }
    expect(() => zonedWallClockToUtcIso("2026-13-45T00:00Z", "UTC")).toThrow(ValidationError);
  });
});

describe("normalizeLocalDateTime input shapes (#33)", () => {
  it("accepts fractional seconds and a space separator, dropping the fraction", () => {
    expect(normalizeLocalDateTime("2026-07-20T10:00:00.000")).toBe("2026-07-20T10:00:00");
    expect(normalizeLocalDateTime("2026-07-20 10:00")).toBe("2026-07-20T10:00:00");
  });
  it("still needs a time for an event: a bare date is rejected", () => {
    expect(() => normalizeLocalDateTime("2026-07-20")).toThrow(ValidationError);
  });
  it("keeps a four-digit year, leading zeros included (Copilot review on #116)", () => {
    expect(normalizeLocalDateTime("0100-07-20T10:00:00")).toBe("0100-07-20T10:00:00");
    // The normalised value parses again, as the ordering checks need.
    expect(() => zonedWallClockToUtcIso("0100-07-20T10:00:00", "UTC")).not.toThrow();
  });
});

describe("zonedWallClockToUtcIso around DST changes (Copilot review on #116)", () => {
  it("maps a skipped midnight to the start of that local day, not the day before", () => {
    // Santiago jumps 00:00 -> 01:00 on 2026-09-06 (UTC-4 -> UTC-3); the day starts at 04:00Z.
    expect(zonedWallClockToUtcIso("2026-09-06", "America/Santiago")).toBe(
      "2026-09-06T04:00:00.000Z",
    );
    // A time inside the gap moves forward by the gap, like the clocks.
    expect(zonedWallClockToUtcIso("2026-09-06T00:30", "America/Santiago")).toBe(
      "2026-09-06T04:30:00.000Z",
    );
    // Havana and Beirut skip midnight too.
    expect(zonedWallClockToUtcIso("2026-03-08", "America/Havana")).toBe("2026-03-08T05:00:00.000Z");
    expect(zonedWallClockToUtcIso("2026-03-29", "Asia/Beirut")).toBe("2026-03-28T22:00:00.000Z");
  });
  it("takes the earlier instant for a time that occurs twice", () => {
    // New York repeats 01:00-02:00 on 2026-11-01; 01:30 EDT is 05:30Z.
    expect(zonedWallClockToUtcIso("2026-11-01T01:30", "America/New_York")).toBe(
      "2026-11-01T05:30:00.000Z",
    );
  });
  it("is exact either side of a change", () => {
    expect(zonedWallClockToUtcIso("2026-09-05T23:59", "America/Santiago")).toBe(
      "2026-09-06T03:59:00.000Z",
    );
    expect(zonedWallClockToUtcIso("2026-09-06T01:00", "America/Santiago")).toBe(
      "2026-09-06T04:00:00.000Z",
    );
    expect(zonedWallClockToUtcIso("2026-11-01T02:00", "America/New_York")).toBe(
      "2026-11-01T07:00:00.000Z",
    );
  });
});

// Southern-hemisphere DST, both ways (#64). Cases first proposed by
// @masyarakatai in #76; the rule is Temporal's "compatible" one (spec §9.5).
describe("zonedWallClockToUtcIso in Sydney and New York across DST (#64)", () => {
  it("handles a normal Sydney time in daylight saving (+11) and in standard time (+10)", () => {
    expect(zonedWallClockToUtcIso("2026-01-15T12:00:00", "Australia/Sydney")).toBe(
      "2026-01-15T01:00:00.000Z",
    );
    expect(zonedWallClockToUtcIso("2026-07-15T12:00:00", "Australia/Sydney")).toBe(
      "2026-07-15T02:00:00.000Z",
    );
  });
  it("moves a time that never occurs (spring forward) on by the gap, like the clocks", () => {
    // Sydney skips 02:00-03:00 on 2026-10-04: 02:30 lands at 03:30 (+11).
    expect(zonedWallClockToUtcIso("2026-10-04T02:30:00", "Australia/Sydney")).toBe(
      "2026-10-03T16:30:00.000Z",
    );
    // New York skips 02:00-03:00 on 2026-03-08: 02:30 lands at 03:30 EDT.
    expect(zonedWallClockToUtcIso("2026-03-08T02:30:00", "America/New_York")).toBe(
      "2026-03-08T07:30:00.000Z",
    );
  });
  it("takes the first occurrence of a time that happens twice (fall back)", () => {
    // Sydney repeats 02:00-03:00 on 2026-04-05: the first 02:30 is still +11.
    expect(zonedWallClockToUtcIso("2026-04-05T02:30:00", "Australia/Sydney")).toBe(
      "2026-04-04T15:30:00.000Z",
    );
    // Just outside the repeated hour, there's only one answer either side.
    expect(zonedWallClockToUtcIso("2026-04-05T01:59:00", "Australia/Sydney")).toBe(
      "2026-04-04T14:59:00.000Z",
    );
    expect(zonedWallClockToUtcIso("2026-04-05T03:00:00", "Australia/Sydney")).toBe(
      "2026-04-04T17:00:00.000Z",
    );
  });
  it("renders both Sydney 02:30s back as 02:30, so a listing shows the wall clock", () => {
    expect(utcToZonedWallClock(Date.parse("2026-04-04T15:30:00Z"), "Australia/Sydney")).toBe(
      "2026-04-05T02:30:00.0000000",
    );
    expect(utcToZonedWallClock(Date.parse("2026-04-04T16:30:00Z"), "Australia/Sydney")).toBe(
      "2026-04-05T02:30:00.0000000",
    );
  });
});

describe("preferTimeZoneHeader", () => {
  it("formats the Prefer header value", () => {
    expect(preferTimeZoneHeader("Australia/Brisbane")).toBe(
      'outlook.timezone="Australia/Brisbane"',
    );
  });
  it("rejects an unknown zone", () => {
    expect(() => preferTimeZoneHeader("Nope/Nope")).toThrow(ValidationError);
  });
});

describe("utcToZonedWallClock", () => {
  it("renders an instant as local wall-clock time, in Graph's form", () => {
    const at = Date.parse("2026-11-01T06:30:00.250Z"); // the second 01:30 in New York
    expect(utcToZonedWallClock(at, "America/New_York")).toBe("2026-11-01T01:30:00.2500000");
    expect(utcToZonedWallClock(Date.parse("2026-07-20T00:00:00Z"), "Australia/Brisbane")).toBe(
      "2026-07-20T10:00:00.0000000",
    );
  });
});
