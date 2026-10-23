import { describe, expect, it, vi } from "vitest";
import { ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { findFreeTime, mergeBusy, subtractIntervals } from "./freebusy.js";

type RequestMock = (
  method: string,
  path: string,
  opts?: { query?: Record<string, unknown>; headers?: Record<string, string> },
) => Promise<{ status: number; data: unknown; headers: Headers }>;

/** An event as calendarView returns it in UTC. */
function ev(start: string, end: string, showAs = "busy", isCancelled = false) {
  return {
    id: `${start}-${showAs}`,
    start: { dateTime: `${start}.0000000`, timeZone: "UTC" },
    end: { dateTime: `${end}.0000000`, timeZone: "UTC" },
    showAs,
    isCancelled,
  };
}

function graphOf(...pages: { value: unknown[]; next?: string }[]) {
  const request = vi.fn<RequestMock>(async () => {
    const p = pages.shift() ?? { value: [] };
    return {
      status: 200,
      data: { value: p.value, ...(p.next ? { "@odata.nextLink": p.next } : {}) },
      headers: new Headers(),
    };
  });
  return { request, graph: { request } as unknown as GraphClient };
}

/** Just the local endpoints, for the tests that aren't about the UTC ones. */
const local = <T extends { startUtc: string; endUtc: string }>(xs: T[]) =>
  xs.map(({ startUtc: _s, endUtc: _e, ...rest }) => rest);

const NEXT = "https://graph.microsoft.com/v1.0/me/calendarView?$skip=100";
const base = {
  startDateTime: "2026-07-20T00:00:00Z",
  endDateTime: "2026-07-21T00:00:00Z",
  timeZone: "UTC",
  minDurationMinutes: 30,
};

describe("interval maths", () => {
  it("merges overlapping and touching busy spans, keeping the strongest status", () => {
    expect(
      mergeBusy([
        { start: 30, end: 40, showAs: "tentative" },
        { start: 0, end: 10, showAs: "busy" },
        { start: 10, end: 20, showAs: "tentative" }, // touches the first
        { start: 35, end: 50, showAs: "oof" },
      ]),
    ).toEqual([
      { start: 0, end: 20, showAs: "busy" },
      { start: 30, end: 50, showAs: "oof" },
    ]);
  });

  it("subtracts busy time from windows", () => {
    expect(
      subtractIntervals(
        [
          { start: 0, end: 100 },
          { start: 200, end: 300 },
        ],
        [
          { start: 10, end: 20 },
          { start: 90, end: 210 },
          { start: 250, end: 260 },
        ],
      ),
    ).toEqual([
      { start: 0, end: 10 },
      { start: 20, end: 90 },
      { start: 210, end: 250 },
      { start: 260, end: 300 },
    ]);
    expect(subtractIntervals([{ start: 0, end: 10 }], [{ start: -5, end: 15 }])).toEqual([]);
  });
});

describe("findFreeTime (#52)", () => {
  it("reads the range in UTC and returns the gaps of at least the minimum length", async () => {
    const { request, graph } = graphOf({
      value: [
        ev("2026-07-20T09:00:00", "2026-07-20T10:00:00"),
        ev("2026-07-20T10:15:00", "2026-07-20T12:00:00"), // leaves a 15-minute gap
        ev("2026-07-20T13:00:00", "2026-07-20T14:00:00", "free"), // doesn't count
        ev("2026-07-20T15:00:00", "2026-07-20T16:00:00", "busy", true), // cancelled
      ],
    });
    const res = await findFreeTime(graph, base);
    const [, path, opts] = request.mock.calls[0]!;
    expect(path).toBe("/me/calendarView");
    expect(opts?.headers?.prefer).toBe('outlook.timezone="UTC"');
    expect(res.freeSlots).toEqual([
      {
        start: "2026-07-20T00:00:00",
        end: "2026-07-20T09:00:00",
        startUtc: "2026-07-20T00:00:00Z",
        endUtc: "2026-07-20T09:00:00Z",
        minutes: 540,
      },
      {
        start: "2026-07-20T12:00:00",
        end: "2026-07-21T00:00:00",
        startUtc: "2026-07-20T12:00:00Z",
        endUtc: "2026-07-21T00:00:00Z",
        minutes: 720,
      },
    ]);
    expect(local(res.busy)).toEqual([
      { start: "2026-07-20T09:00:00", end: "2026-07-20T10:00:00", showAs: "busy" },
      { start: "2026-07-20T10:15:00", end: "2026-07-20T12:00:00", showAs: "busy" },
    ]);
  });

  it("counts tentative as busy unless told not to", async () => {
    const tentative = () =>
      graphOf({ value: [ev("2026-07-20T09:00:00", "2026-07-20T17:00:00", "tentative")] }).graph;
    expect((await findFreeTime(tentative(), base)).freeSlots.map((s) => s.minutes)).toEqual([
      540, 420,
    ]);
    expect(
      local((await findFreeTime(tentative(), { ...base, tentativeIsBusy: false })).freeSlots),
    ).toEqual([{ start: "2026-07-20T00:00:00", end: "2026-07-21T00:00:00", minutes: 1440 }]);
  });

  it("looks only within daily hours, per local day, moving with a DST change", async () => {
    // Sydney springs forward on Sunday 2026-10-04 (+10 -> +11); 09:00-17:00 local each day.
    const { graph } = graphOf({ value: [ev("2026-10-04T00:00:00", "2026-10-04T01:00:00")] }); // 11:00-12:00 local
    const res = await findFreeTime(graph, {
      startDateTime: "2026-10-03",
      endDateTime: "2026-10-05",
      timeZone: "Australia/Sydney",
      minDurationMinutes: 60,
      withinHours: { start: "09:00", end: "17:00" },
    });
    expect(local(res.freeSlots)).toEqual([
      { start: "2026-10-03T09:00:00", end: "2026-10-03T17:00:00", minutes: 480 },
      { start: "2026-10-04T09:00:00", end: "2026-10-04T11:00:00", minutes: 120 },
      { start: "2026-10-04T12:00:00", end: "2026-10-04T17:00:00", minutes: 300 },
    ]);
    expect(res.busy).toEqual([
      {
        start: "2026-10-04T11:00:00",
        end: "2026-10-04T12:00:00",
        startUtc: "2026-10-04T00:00:00Z",
        endUtc: "2026-10-04T01:00:00Z",
        showAs: "busy",
      },
    ]);
    // 09:00 is +10 before the change and +11 after.
    expect(res.freeSlots.map((s) => s.startUtc)).toEqual([
      "2026-10-02T23:00:00Z",
      "2026-10-03T22:00:00Z",
      "2026-10-04T01:00:00Z",
    ]);
  });

  it("keeps daily hours that fall in a skipped hour inside those hours", async () => {
    // Sydney skips 02:00-03:00 on 2026-10-04 (the jump is at 16:00Z on the 3rd).
    const day = {
      startDateTime: "2026-10-04",
      endDateTime: "2026-10-05",
      timeZone: "Australia/Sydney",
    };
    const within = (start: string, end: string) =>
      findFreeTime(graphOf().graph, {
        ...day,
        minDurationMinutes: 5,
        withinHours: { start, end },
      }).then((r) => r.freeSlots.map((s) => [s.start, s.end, s.startUtc]));
    // Both bounds skipped: those hours never happen that day.
    expect(await within("02:00", "02:30")).toEqual([]);
    // One bound skipped: the window starts, or ends, at the jump.
    expect(await within("02:30", "04:00")).toEqual([
      ["2026-10-04T03:00:00", "2026-10-04T04:00:00", "2026-10-03T16:00:00Z"],
    ]);
    expect(await within("01:00", "02:30")).toEqual([
      ["2026-10-04T01:00:00", "2026-10-04T03:00:00", "2026-10-03T15:00:00Z"],
    ]);
  });

  it("treats the hour repeated as clocks go back as two separate windows", async () => {
    // New York repeats 01:00-02:00 on 2026-11-01 (05:00-06:00Z, then 06:00-07:00Z).
    const ny = (start: string, end: string, minDurationMinutes = 5) =>
      findFreeTime(graphOf().graph, {
        startDateTime: "2026-11-01",
        endDateTime: "2026-11-02",
        timeZone: "America/New_York",
        minDurationMinutes,
        withinHours: { start, end },
      }).then((r) => r.freeSlots.map((s) => [s.startUtc, s.endUtc, s.minutes]));
    // Both occurrences of 01:30-02:00, never the 01:00-01:30 between them.
    expect(await ny("01:30", "02:00")).toEqual([
      ["2026-11-01T05:30:00Z", "2026-11-01T06:00:00Z", 30],
      ["2026-11-01T06:30:00Z", "2026-11-01T07:00:00Z", 30],
    ]);
    // Neither is an hour long.
    expect(await ny("01:30", "02:00", 60)).toEqual([]);
    // Hours spanning the change stay one window, two hours longer in real time.
    expect(await ny("00:00", "03:00")).toEqual([
      ["2026-11-01T04:00:00Z", "2026-11-01T08:00:00Z", 240],
    ]);
  });

  it("gives the exact instant where the local time repeats as clocks go back", async () => {
    // New York falls back at 06:00Z on 2026-11-01: 01:00-02:00 local happens twice.
    const { graph } = graphOf({ value: [ev("2026-11-01T05:00:00", "2026-11-01T06:00:00")] });
    const res = await findFreeTime(graph, {
      startDateTime: "2026-11-01T05:00:00Z",
      endDateTime: "2026-11-01T06:30:00Z",
      timeZone: "America/New_York",
      minDurationMinutes: 30,
    });
    // Local 01:00-01:30 is ambiguous; the UTC instants say it's the second one.
    expect(res.freeSlots).toEqual([
      {
        start: "2026-11-01T01:00:00",
        end: "2026-11-01T01:30:00",
        startUtc: "2026-11-01T06:00:00Z",
        endUtc: "2026-11-01T06:30:00Z",
        minutes: 30,
      },
    ]);
    expect(res.busy).toEqual([
      {
        start: "2026-11-01T01:00:00",
        end: "2026-11-01T01:00:00",
        startUtc: "2026-11-01T05:00:00Z",
        endUtc: "2026-11-01T06:00:00Z",
        showAs: "busy",
      },
    ]);
  });

  it("reads an all-day event as its local days, not as UTC (#158)", async () => {
    // As Graph returned it live: dates labelled with the requested zone (UTC).
    const allDay = { ...ev("2026-10-23T00:00:00", "2026-10-24T00:00:00"), isAllDay: true };
    const la = await findFreeTime(graphOf({ value: [allDay] }).graph, {
      startDateTime: "2026-10-23",
      endDateTime: "2026-10-25",
      timeZone: "America/Los_Angeles",
      minDurationMinutes: 30,
    });
    // Friday is busy all day, local time; Saturday is free.
    expect(la.busy).toEqual([
      {
        start: "2026-10-23T00:00:00",
        end: "2026-10-24T00:00:00",
        startUtc: "2026-10-23T07:00:00Z",
        endUtc: "2026-10-24T07:00:00Z",
        showAs: "busy",
      },
    ]);
    expect(local(la.freeSlots)).toEqual([
      { start: "2026-10-24T00:00:00", end: "2026-10-25T00:00:00", minutes: 1440 },
    ]);
    // East of UTC too: early Friday morning in Sydney is still that all-day event.
    const sydney = await findFreeTime(graphOf({ value: [allDay] }).graph, {
      startDateTime: "2026-10-23T00:00:00",
      endDateTime: "2026-10-23T09:00:00",
      timeZone: "Australia/Sydney",
      minDurationMinutes: 30,
    });
    expect(sydney.freeSlots).toEqual([]);
  });

  it("gives an all-day event on a clock-change day its real length (#158)", async () => {
    // Los Angeles falls back on Sunday 2026-11-01: that day is 25 hours long.
    const allDay = { ...ev("2026-11-01T00:00:00", "2026-11-02T00:00:00"), isAllDay: true };
    const res = await findFreeTime(graphOf({ value: [allDay] }).graph, {
      startDateTime: "2026-10-31",
      endDateTime: "2026-11-03",
      timeZone: "America/Los_Angeles",
      minDurationMinutes: 30,
    });
    // Local midnight to local midnight: -07:00 before the change, -08:00 after.
    expect(res.busy.map((b) => [b.startUtc, b.endUtc])).toEqual([
      ["2026-11-01T07:00:00Z", "2026-11-02T08:00:00Z"],
    ]);
    expect(res.freeSlots.map((s) => [s.start, s.end])).toEqual([
      ["2026-10-31T00:00:00", "2026-11-01T00:00:00"],
      ["2026-11-02T00:00:00", "2026-11-03T00:00:00"],
    ]);
  });

  it("doesn't let an event outside the range change the status inside it", async () => {
    // Away 07:00-08:00 (outside a 09:00-17:00 lookup) touches busy 08:00-10:00.
    const { graph } = graphOf({
      value: [
        ev("2026-07-20T07:00:00", "2026-07-20T08:00:00", "oof"),
        ev("2026-07-20T08:00:00", "2026-07-20T10:00:00"),
      ],
    });
    const res = await findFreeTime(graph, {
      ...base,
      startDateTime: "2026-07-20T09:00:00Z",
      endDateTime: "2026-07-20T17:00:00Z",
    });
    expect(local(res.busy)).toEqual([
      { start: "2026-07-20T09:00:00", end: "2026-07-20T10:00:00", showAs: "busy" },
    ]);
  });

  it("reads a day past each end of the range, where all-day events can sit (#158)", async () => {
    const { request, graph } = graphOf();
    await findFreeTime(graph, base);
    const query = request.mock.calls[0]![2]?.query;
    expect([query?.startDateTime, query?.endDateTime]).toEqual([
      "2026-07-19T00:00:00.000Z",
      "2026-07-22T00:00:00.000Z",
    ]);
  });

  it("clips busy time to the range and reads every page", async () => {
    const { request, graph } = graphOf(
      { value: [ev("2026-07-19T22:00:00", "2026-07-20T02:00:00")], next: NEXT },
      { value: [ev("2026-07-20T23:00:00", "2026-07-21T03:00:00", "oof")] },
    );
    const res = await findFreeTime(graph, base);
    expect(request).toHaveBeenCalledTimes(2);
    expect(local(res.busy)).toEqual([
      { start: "2026-07-20T00:00:00", end: "2026-07-20T02:00:00", showAs: "busy" },
      { start: "2026-07-20T23:00:00", end: "2026-07-21T00:00:00", showAs: "oof" },
    ]);
    expect(local(res.freeSlots)).toEqual([
      { start: "2026-07-20T02:00:00", end: "2026-07-20T23:00:00", minutes: 1260 },
    ]);
  });

  it("refuses a range it can't read in full, rather than invent free time", async () => {
    const pages = Array.from({ length: 10 }, () => ({ value: [], next: NEXT }));
    await expect(findFreeTime(graphOf(...pages).graph, base)).rejects.toThrow(/too many events/);
    // A next page that can't be followed, likewise.
    await expect(
      findFreeTime(graphOf({ value: [], next: "https://elsewhere.example/next" }).graph, base),
    ).rejects.toThrow(/too many events/);
  });

  it("refuses, rather than call it free, when a busy event's times can't be read", async () => {
    const broken = { ...ev("2026-07-20T09:00:00", "2026-07-20T10:00:00"), end: null };
    await expect(findFreeTime(graphOf({ value: [broken] }).graph, base)).rejects.toThrow(
      /couldn't be read/,
    );
    // A free or cancelled one doesn't matter.
    const freeBroken = { ...ev("2026-07-20T09:00:00", "2026-07-20T10:00:00", "free"), end: null };
    const res = await findFreeTime(graphOf({ value: [freeBroken] }).graph, base);
    expect(res.freeSlots).toHaveLength(1);
  });

  it("counts only whole free minutes", async () => {
    const res = await findFreeTime(graphOf().graph, {
      ...base,
      startDateTime: "2026-07-20T09:00:00Z",
      endDateTime: "2026-07-20T09:30:40Z",
    });
    expect(res.freeSlots.map((s) => s.minutes)).toEqual([30]);
  });

  it("keeps slots to whole seconds inside the free time", async () => {
    // Busy until 10:00:00.5; free from then until 10:30:00.2.
    const { graph } = graphOf({
      value: [
        {
          ...ev("2026-07-20T09:00:00", "2026-07-20T10:00:00"),
          end: { dateTime: "2026-07-20T10:00:00.5000000", timeZone: "UTC" },
        },
      ],
    });
    const res = await findFreeTime(graph, {
      ...base,
      startDateTime: "2026-07-20T09:00:00Z",
      endDateTime: "2026-07-20T10:30:00.200Z",
      minDurationMinutes: 30,
    });
    // 10:00:01-10:30:00 is under 30 minutes once rounded inward.
    expect(res.freeSlots).toEqual([]);
    const shorter = await findFreeTime(graphOf({ value: [] }).graph, {
      ...base,
      startDateTime: "2026-07-20T10:00:00.500Z",
      endDateTime: "2026-07-20T10:40:00.200Z",
      minDurationMinutes: 30,
    });
    expect(shorter.freeSlots.map((s) => [s.startUtc, s.endUtc])).toEqual([
      ["2026-07-20T10:00:01Z", "2026-07-20T10:40:00Z"],
    ]);
  });

  it("checks its inputs before calling Graph", async () => {
    const { request, graph } = graphOf();
    const bad = [
      { ...base, endDateTime: base.startDateTime },
      { ...base, endDateTime: "2026-10-01T00:00:00Z" }, // over 62 days
      { ...base, minDurationMinutes: 0 },
      { ...base, withinHours: { start: "9:00", end: "17:00" } },
      { ...base, withinHours: { start: "17:00", end: "09:00" } },
      { ...base, timeZone: "Nowhere/Nope" },
    ];
    for (const input of bad) {
      await expect(findFreeTime(graph, input)).rejects.toBeInstanceOf(ValidationError);
    }
    expect(request).not.toHaveBeenCalled();
  });

  it("reads another calendar by its id", async () => {
    const { request, graph } = graphOf();
    await findFreeTime(graph, { ...base, calendarId: "cal/2" });
    expect(request.mock.calls[0]![1]).toBe("/me/calendars/cal%2F2/calendarView");
  });
});
