import { describe, expect, it } from "vitest";
import { ValidationError } from "./errors.js";
import {
  isValidTimeZone,
  normalizeLocalDateTime,
  preferTimeZoneHeader,
  shiftLocalDateTime,
  toGraphDateTime,
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

describe("shiftLocalDateTime — travel blocks", () => {
  it("subtracts and adds minutes on the wall clock", () => {
    expect(shiftLocalDateTime("2026-07-20T12:00:00", -15)).toBe("2026-07-20T11:45:00");
    expect(shiftLocalDateTime("2026-07-20T12:00:00", 15)).toBe("2026-07-20T12:15:00");
  });
  it("rolls over day boundaries", () => {
    expect(shiftLocalDateTime("2026-07-20T23:50:00", 15)).toBe("2026-07-21T00:05:00");
    expect(shiftLocalDateTime("2026-07-20T00:05:00", -15)).toBe("2026-07-19T23:50:00");
  });
  it("rolls over month/year boundaries", () => {
    expect(shiftLocalDateTime("2026-12-31T23:45:00", 30)).toBe("2027-01-01T00:15:00");
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
