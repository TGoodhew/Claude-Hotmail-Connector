import { describe, expect, it } from "vitest";
import { ValidationError } from "./errors.js";
import {
  hasExplicitOffset,
  isValidTimeZone,
  normalizeLocalDateTime,
  preferTimeZoneHeader,
  toGraphDateTime,
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
  it("passes a date-only value through untouched", () => {
    expect(zonedWallClockToUtcIso("2026-07-20", "Australia/Brisbane")).toBe("2026-07-20");
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
