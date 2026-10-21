import { describe, expect, it, vi } from "vitest";
import { ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { listEvents } from "./calendar.js";

type GetAllPagesMock = (
  path: string,
  opts?: { query?: Record<string, unknown>; headers?: Record<string, string> },
  pageOpts?: { maxItems?: number },
) => Promise<unknown[]>;

function graphWith(getAllPages: GetAllPagesMock): GraphClient {
  return { getAllPages } as unknown as GraphClient;
}

const validRange = {
  startDateTime: "2026-07-20T00:00:00+10:00",
  endDateTime: "2026-07-21T00:00:00+10:00",
  timeZone: "Australia/Brisbane",
  limit: 25,
};

describe("listEvents", () => {
  it("calls calendarView with the Prefer header and window params", async () => {
    const getAllPages = vi.fn<GetAllPagesMock>(async () => []);
    await listEvents(graphWith(getAllPages), validRange);

    const [path, opts, pageOpts] = getAllPages.mock.calls[0]!;
    expect(path).toBe("/me/calendarView");
    expect(opts?.headers?.prefer).toBe('outlook.timezone="Australia/Brisbane"');
    expect(opts?.query?.startDateTime).toBe(validRange.startDateTime);
    expect(opts?.query?.endDateTime).toBe(validRange.endDateTime);
    expect(opts?.query?.$orderby).toBe("start/dateTime");
    expect(pageOpts?.maxItems).toBe(25);
  });

  it("shapes events", async () => {
    const getAllPages = vi.fn<GetAllPagesMock>(async () => [
      {
        id: "e1",
        subject: "Macadamias Australia — visit",
        start: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
        end: { dateTime: "2026-07-20T11:00:00.0000000", timeZone: "Australia/Brisbane" },
        location: { displayName: "Bundaberg QLD" },
        isAllDay: false,
        webLink: "https://outlook/e1",
      },
    ]);
    const events = await listEvents(graphWith(getAllPages), validRange);
    expect(events).toEqual([
      {
        id: "e1",
        subject: "Macadamias Australia — visit",
        start: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
        end: { dateTime: "2026-07-20T11:00:00.0000000", timeZone: "Australia/Brisbane" },
        location: "Bundaberg QLD",
        isAllDay: false,
        webLink: "https://outlook/e1",
      },
    ]);
  });

  it("rejects an invalid time zone", async () => {
    const getAllPages = vi.fn<GetAllPagesMock>(async () => []);
    await expect(
      listEvents(graphWith(getAllPages), { ...validRange, timeZone: "Nowhere/Nope" }),
    ).rejects.toBeInstanceOf(ValidationError);
  });

  it("rejects invalid range dates", async () => {
    const getAllPages = vi.fn<GetAllPagesMock>(async () => []);
    await expect(
      listEvents(graphWith(getAllPages), { ...validRange, startDateTime: "nope" }),
    ).rejects.toBeInstanceOf(ValidationError);
  });

  it("pins a date-only window to the zone's midnights (#33)", async () => {
    const getAllPages = vi.fn<GetAllPagesMock>(async () => []);
    await listEvents(graphWith(getAllPages), {
      ...validRange,
      startDateTime: "2026-07-20",
      endDateTime: "2026-07-21",
    });
    expect(getAllPages.mock.calls[0]![1]?.query).toMatchObject({
      startDateTime: "2026-07-19T14:00:00.000Z",
      endDateTime: "2026-07-20T14:00:00.000Z",
    });
  });

  it("rejects a window that ends at or before it starts, without calling Graph (#33)", async () => {
    const getAllPages = vi.fn<GetAllPagesMock>(async () => []);
    for (const endDateTime of ["2026-07-19T23:00:00+10:00", validRange.startDateTime]) {
      await expect(
        listEvents(graphWith(getAllPages), { ...validRange, endDateTime }),
      ).rejects.toThrow(/endDateTime must be after startDateTime/);
    }
    expect(getAllPages).not.toHaveBeenCalled();
  });

  it("orders sub-second bounds correctly, local or mixed with an offset (Copilot review on #116)", async () => {
    const getAllPages = vi.fn<GetAllPagesMock>(async () => []);
    const run = (startDateTime: string, endDateTime: string) =>
      listEvents(graphWith(getAllPages), { ...validRange, startDateTime, endDateTime });

    // An 800 ms local window is valid, and both bounds keep their milliseconds.
    await run("2026-07-20T10:00:00.100", "2026-07-20T10:00:00.900");
    expect(getAllPages.mock.calls[0]![1]?.query).toMatchObject({
      startDateTime: "2026-07-20T00:00:00.100Z",
      endDateTime: "2026-07-20T00:00:00.900Z",
    });
    // A local start after an offset end, by 800 ms, is reversed.
    await expect(run("2026-07-20T10:00:00.900", "2026-07-20T10:00:00.100+10:00")).rejects.toThrow(
      /endDateTime must be after startDateTime/,
    );
    // And the mixed case the right way round is fine.
    await run("2026-07-20T10:00:00.100", "2026-07-20T10:00:00.900+10:00");
    expect(getAllPages).toHaveBeenCalledTimes(2);
  });
});
