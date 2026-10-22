import { createHash } from "node:crypto";
import { describe, expect, it, vi } from "vitest";
import { GraphError, ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import {
  assertWritableEventPayload,
  buildRecurrence,
  cancelEvent,
  createEvent,
  type CreateEventInput,
  type RecurrenceInput,
  updateEvent,
  WRITABLE_EVENT_FIELDS,
} from "./calendar.js";

type PostMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;
type PatchMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;
type GetMock = (path: string, opts?: unknown) => Promise<unknown>;
type ReqOpts = { query: Record<string, string>; headers?: Record<string, string> };

function graphWith(fns: { post?: PostMock; patch?: PatchMock; get?: GetMock }): GraphClient {
  // update_event reads the current values first; default to an empty event.
  return {
    get: async () => ({ id: "evt-1", "@odata.etag": 'W/"t1"' }),
    ...fns,
  } as unknown as GraphClient;
}

const createdReply = {
  id: "evt-1",
  subject: "Macadamias Australia — visit",
  start: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
  end: { dateTime: "2026-07-20T11:00:00.0000000", timeZone: "Australia/Brisbane" },
  location: { displayName: "Bundaberg QLD" },
  webLink: "https://outlook/evt-1",
};

describe("createEvent", () => {
  it("sends local wall-clock + IANA time zone (no UTC shift) and a transactionId", async () => {
    const post = vi.fn<PostMock>(async () => createdReply);
    const result = await createEvent(
      graphWith({ post }),
      {
        subject: "Macadamias Australia — visit",
        start: { dateTime: "2026-07-20T10:00:00", timeZone: "Australia/Brisbane" },
        end: { dateTime: "2026-07-20T11:00:00", timeZone: "Australia/Brisbane" },
        location: "Bundaberg QLD",
      },
      () => "fixed-txn-id",
    );

    const [path, body] = post.mock.calls[0]!;
    expect(path).toBe("/me/events");
    const payload = body as Record<string, unknown>;
    // The critical assertion: the wall-clock time is preserved, not shifted.
    expect(payload.start).toEqual({
      dateTime: "2026-07-20T10:00:00",
      timeZone: "Australia/Brisbane",
    });
    expect(payload.end).toEqual({
      dateTime: "2026-07-20T11:00:00",
      timeZone: "Australia/Brisbane",
    });
    // The dateTime value itself must not carry a UTC 'Z'.
    expect((payload.start as { dateTime: string }).dateTime).not.toContain("Z");
    expect(payload.transactionId).toBe("fixed-txn-id");
    expect(payload.location).toEqual({ displayName: "Bundaberg QLD" });
    expect(result.id).toBe("evt-1");
    expect(result.transactionId).toBe("fixed-txn-id");
  });

  it("creates in another calendar when given its calendarId (#49)", async () => {
    const post = vi.fn<PostMock>(async () => createdReply);
    await createEvent(graphWith({ post }), {
      subject: "Trip",
      start: { dateTime: "2026-07-20T10:00:00", timeZone: "Australia/Brisbane" },
      end: { dateTime: "2026-07-20T11:00:00", timeZone: "Australia/Brisbane" },
      calendarId: "cal/2",
    });
    expect(post.mock.calls[0]![0]).toBe("/me/calendars/cal%2F2/events");
    // The calendar is in the path, never in the payload.
    expect(post.mock.calls[0]![1]).not.toHaveProperty("calendarId");
  });

  it("updates within another calendar when given its calendarId (#49)", async () => {
    const get = vi.fn<GetMock>(async () => ({ id: "evt-1", "@odata.etag": 'W/"t1"' }));
    const patch = vi.fn<PatchMock>(async () => createdReply);
    await updateEvent(graphWith({ get, patch }), {
      id: "evt-1",
      calendarId: "cal/2",
      subject: "X",
    });
    expect(get.mock.calls[0]![0]).toBe("/me/calendars/cal%2F2/events/evt-1");
    expect(patch.mock.calls[0]![0]).toBe("/me/calendars/cal%2F2/events/evt-1");
  });

  it("uses a caller-supplied transactionId for idempotency", async () => {
    const post = vi.fn<PostMock>(async () => createdReply);
    await createEvent(graphWith({ post }), {
      subject: "X",
      start: { dateTime: "2026-07-20T10:00:00", timeZone: "Australia/Brisbane" },
      end: { dateTime: "2026-07-20T11:00:00", timeZone: "Australia/Brisbane" },
      transactionId: "my-key",
    });
    expect((post.mock.calls[0]![1] as Record<string, unknown>).transactionId).toBe("my-key");
  });

  it("rejects a UTC 'Z' dateTime (would land at the wrong local time)", async () => {
    const post = vi.fn<PostMock>(async () => createdReply);
    await expect(
      createEvent(graphWith({ post }), {
        subject: "X",
        start: { dateTime: "2026-07-20T10:00:00Z", timeZone: "Australia/Brisbane" },
        end: { dateTime: "2026-07-20T11:00:00", timeZone: "Australia/Brisbane" },
      }),
    ).rejects.toBeInstanceOf(ValidationError);
    expect(post).not.toHaveBeenCalled();
  });
});

describe("all-day, free/busy and recurring creation (#51)", () => {
  const posted = async (input: Partial<CreateEventInput>, reply: object = createdReply) => {
    const post = vi.fn<PostMock>(async () => reply);
    const result = await createEvent(graphWith({ post }), {
      subject: "Trip",
      start: { dateTime: "2026-07-20T10:00:00", timeZone: "Australia/Brisbane" },
      end: { dateTime: "2026-07-20T11:00:00", timeZone: "Australia/Brisbane" },
      ...input,
    });
    return { payload: post.mock.calls[0]![1] as Record<string, unknown>, result, post };
  };

  it("writes an all-day event as midnight to midnight in its own zone, never shifted", async () => {
    for (const timeZone of ["Australia/Brisbane", "America/New_York", "Asia/Kolkata", "UTC"]) {
      const { payload } = await posted({
        isAllDay: true,
        start: { dateTime: "2026-07-20", timeZone },
        end: { dateTime: "2026-07-22", timeZone }, // two days: the 20th and 21st
      });
      expect(payload).toMatchObject({
        isAllDay: true,
        start: { dateTime: "2026-07-20T00:00:00", timeZone },
        end: { dateTime: "2026-07-22T00:00:00", timeZone },
      });
    }
    // Midnight given as a time is fine too.
    const { payload } = await posted({
      isAllDay: true,
      start: { dateTime: "2026-07-20T00:00", timeZone: "Europe/London" },
      end: { dateTime: "2026-07-21T00:00:00", timeZone: "Europe/London" },
    });
    expect(payload.start).toEqual({ dateTime: "2026-07-20T00:00:00", timeZone: "Europe/London" });
    // So is an all-zero fraction of a second.
    const zero = await posted({
      isAllDay: true,
      start: { dateTime: "2026-07-20T00:00:00.000", timeZone: "Europe/London" },
      end: { dateTime: "2026-07-21", timeZone: "Europe/London" },
    });
    expect(zero.payload.start).toEqual({
      dateTime: "2026-07-20T00:00:00",
      timeZone: "Europe/London",
    });
  });

  it("refuses an all-day event that isn't whole days in one zone, before writing", async () => {
    const tz = "Australia/Brisbane";
    const cases: [Partial<CreateEventInput>, RegExp][] = [
      [
        {
          start: { dateTime: "2026-07-20T09:00:00", timeZone: tz },
          end: { dateTime: "2026-07-21", timeZone: tz },
        },
        /start must be a date/,
      ],
      [
        {
          start: { dateTime: "2026-07-20", timeZone: tz },
          end: { dateTime: "2026-07-21T12:00", timeZone: tz },
        },
        /end must be a date/,
      ],
      [
        {
          start: { dateTime: "2026-07-20", timeZone: tz },
          end: { dateTime: "2026-07-21", timeZone: "UTC" },
        },
        /same time zone/,
      ],
      [
        {
          start: { dateTime: "2026-07-20", timeZone: tz },
          end: { dateTime: "2026-07-20", timeZone: tz },
        },
        /end must be after its start/,
      ],
      [
        // Half a second past midnight isn't midnight (Copilot review on #134).
        {
          start: { dateTime: "2026-07-20T00:00:00.5", timeZone: tz },
          end: { dateTime: "2026-07-21", timeZone: tz },
        },
        /start must be a date/,
      ],
      [
        {
          start: { dateTime: "2026-07-20", timeZone: tz },
          end: { dateTime: "2026-07-21T00:00:00.001", timeZone: tz },
        },
        /end must be a date/,
      ],
    ];
    for (const [input, error] of cases) {
      const post = vi.fn<PostMock>(async () => createdReply);
      await expect(
        createEvent(graphWith({ post }), {
          subject: "Holiday",
          isAllDay: true,
          ...input,
        } as CreateEventInput),
      ).rejects.toThrow(error);
      expect(post).not.toHaveBeenCalled();
    }
  });

  it("sets showAs, e.g. free for a travel block", async () => {
    const { payload } = await posted({ showAs: "free" });
    expect(payload.showAs).toBe("free");
    const plain = await posted({});
    expect(plain.payload).not.toHaveProperty("showAs");
    expect(plain.payload).not.toHaveProperty("isAllDay");
    expect(plain.payload).not.toHaveProperty("recurrence");
  });

  it("writes a recurring series in the event's zone, and echoes the new fields", async () => {
    const recurrence: RecurrenceInput = {
      pattern: { type: "weekly", interval: 1, daysOfWeek: ["monday", "wednesday"] },
      range: { type: "numbered", startDate: "2026-07-20", numberOfOccurrences: 10 },
    };
    const reply = {
      ...createdReply,
      isAllDay: false,
      showAs: "busy",
      recurrence: { from: "graph" },
    };
    const { payload, result } = await posted({ recurrence }, reply);
    expect(payload.recurrence).toEqual({
      pattern: { type: "weekly", interval: 1, daysOfWeek: ["monday", "wednesday"] },
      range: {
        type: "numbered",
        startDate: "2026-07-20",
        numberOfOccurrences: 10,
        recurrenceTimeZone: "Australia/Brisbane",
      },
    });
    expect(result).toMatchObject({
      isAllDay: false,
      showAs: "busy",
      recurrence: { from: "graph" },
    });
    const plain = await posted({}, createdReply);
    expect(plain.result).toMatchObject({ isAllDay: false, showAs: null, recurrence: null });
  });

  it("gives a recurring or all-day create its own idempotency key, keeping a plain one's", async () => {
    const key = async (input: Partial<CreateEventInput>) =>
      (await posted(input)).payload.transactionId as string;
    const plain = await key({});
    expect(await key({})).toBe(plain); // stable
    const daily: RecurrenceInput = {
      pattern: { type: "daily", interval: 1 },
      range: { type: "noEnd", startDate: "2026-07-20" },
    };
    expect(await key({ recurrence: daily })).not.toBe(plain);
    const allDay = {
      isAllDay: true,
      start: { dateTime: "2026-07-20T00:00:00", timeZone: "Australia/Brisbane" },
      end: { dateTime: "2026-07-21T00:00:00", timeZone: "Australia/Brisbane" },
    };
    expect(await key(allDay)).not.toBe(await key({ ...allDay, isAllDay: undefined }));
  });

  it("keys equivalent inputs alike, so a retry phrased differently doesn't duplicate (Copilot review on #134)", async () => {
    const key = async (input: Partial<CreateEventInput>) =>
      (await posted(input)).payload.transactionId as string;
    const daily: RecurrenceInput = {
      pattern: { type: "daily", interval: 1 },
      range: { type: "noEnd", startDate: "2026-07-20" },
    };
    // Fields the pattern and range types ignore are dropped before keying.
    const padded: RecurrenceInput = {
      pattern: { type: "daily", interval: 1, dayOfMonth: 3, daysOfWeek: ["monday"] },
      range: { type: "noEnd", startDate: "2026-07-20", endDate: "2026-12-31" },
    };
    expect(await key({ recurrence: padded })).toBe(await key({ recurrence: daily }));
    // An all-day date and its midnight are the same event.
    const tz = "Australia/Brisbane";
    const asDates = {
      isAllDay: true,
      start: { dateTime: "2026-07-20", timeZone: tz },
      end: { dateTime: "2026-07-21", timeZone: tz },
    };
    const asMidnights = {
      isAllDay: true,
      start: { dateTime: "2026-07-20T00:00:00", timeZone: tz },
      end: { dateTime: "2026-07-21T00:00", timeZone: tz },
    };
    expect(await key(asDates)).toBe(await key(asMidnights));
    // The same weekdays in another order, or repeated, are the same series.
    const weekly = (daysOfWeek: RecurrenceInput["pattern"]["daysOfWeek"]): RecurrenceInput => ({
      pattern: { type: "weekly", interval: 1, daysOfWeek },
      range: { type: "noEnd", startDate: "2026-07-20" },
    });
    const monWed = await key({ recurrence: weekly(["monday", "wednesday"]) });
    expect(await key({ recurrence: weekly(["wednesday", "monday"]) })).toBe(monWed);
    expect(await key({ recurrence: weekly(["wednesday", "monday", "monday"]) })).toBe(monWed);
    const { payload } = await posted({ recurrence: weekly(["wednesday", "monday", "monday"]) });
    expect((payload.recurrence as { pattern: unknown }).pattern).toMatchObject({
      daysOfWeek: ["monday", "wednesday"],
    });
    // Sunday as the first day of the week is Graph's default: implicit either way.
    const sundayFirst = weekly(["monday", "wednesday"]);
    sundayFirst.pattern.firstDayOfWeek = "sunday";
    expect(await key({ recurrence: sundayFirst })).toBe(monWed);
    const sundayPayload = (await posted({ recurrence: sundayFirst })).payload;
    expect((sundayPayload.recurrence as { pattern: object }).pattern).not.toHaveProperty(
      "firstDayOfWeek",
    );
    const mondayFirst = weekly(["monday", "wednesday"]);
    mondayFirst.pattern.firstDayOfWeek = "monday";
    expect(await key({ recurrence: mondayFirst })).not.toBe(monWed); // a real difference
    // A series keys on its written times, whatever form they were given in.
    const at = (start: string, end: string): Partial<CreateEventInput> => ({
      start: { dateTime: start, timeZone: tz },
      end: { dateTime: end, timeZone: tz },
      recurrence: weekly(["monday"]),
    });
    expect(await key(at("2026-07-20T10:00", "2026-07-20T11:00"))).toBe(
      await key(at("2026-07-20T10:00:00", "2026-07-20T11:00:00.000")),
    );
  });

  it("keys a non-default showAs, and treats busy as the default (Copilot review on #134)", async () => {
    const key = async (input: Partial<CreateEventInput>) =>
      (await posted(input)).payload.transactionId as string;
    const plain = await key({});
    expect(await key({ showAs: "busy" })).toBe(plain); // busy is what an omitted showAs means
    const free = await key({ showAs: "free" });
    expect(free).not.toBe(plain); // a free block isn't a retry of a busy event
    expect(await key({ showAs: "tentative" })).not.toBe(free);
  });
});

describe("create_event polish from the live run (#135)", () => {
  const tz = "America/Los_Angeles";
  const allDay: Partial<CreateEventInput> = {
    isAllDay: true,
    start: { dateTime: "2026-11-20", timeZone: tz },
    end: { dateTime: "2026-11-21", timeZone: tz },
  };
  const create = async (
    input: Partial<CreateEventInput>,
    reply: object = createdReply,
    now = () => Date.parse("2026-10-22T15:00:00Z"),
  ) => {
    const post = vi.fn<PostMock>(async () => reply);
    const result = await createEvent(
      graphWith({ post }),
      {
        subject: "Trip",
        start: { dateTime: "2026-11-23T09:00:00", timeZone: tz },
        end: { dateTime: "2026-11-23T09:30:00", timeZone: tz },
        ...input,
      },
      undefined,
      now,
    );
    return { payload: post.mock.calls[0]![1] as Record<string, unknown>, result };
  };

  it("shows an all-day event as free unless told otherwise, as Outlook does", async () => {
    expect((await create(allDay)).payload.showAs).toBe("free");
    expect((await create({ ...allDay, showAs: "busy" })).payload.showAs).toBe("busy");
    expect((await create({})).payload).not.toHaveProperty("showAs"); // timed: Graph's busy
    // The implicit free keys like an explicit one, and busy differs.
    const key = async (input: Partial<CreateEventInput>) =>
      (await create(input)).payload.transactionId as string;
    expect(await key({ ...allDay, showAs: "free" })).toBe(await key(allDay));
    expect(await key({ ...allDay, showAs: "busy" })).not.toBe(await key(allDay));
  });

  it("keeps the key encoding of earlier versions, so a retry across an upgrade matches (Copilot review on #140)", async () => {
    // The material a key was hashed from before all-day events defaulted to
    // free: subject, written times, location, the all-day flag, and a showAs
    // token unless busy.
    const legacyKey = (showAsToken?: string) =>
      createHash("sha256")
        .update(
          [
            "Trip",
            "2026-11-20T00:00:00",
            tz,
            "2026-11-21T00:00:00",
            tz,
            "",
            "isAllDay",
            ...(showAsToken ? [showAsToken] : []),
          ].join(String.fromCharCode(0)),
        )
        .digest("hex")
        .slice(0, 36);
    const key = async (input: Partial<CreateEventInput>) =>
      (await create(input)).payload.transactionId as string;
    expect(await key({ ...allDay, showAs: "busy" })).toBe(legacyKey());
    expect(await key({ ...allDay, showAs: "free" })).toBe(legacyKey("showAs:free"));
    // The new implicit free is a free event: it keys as one, not as the old busy.
    expect(await key(allDay)).toBe(legacyKey("showAs:free"));
  });

  it("says when Graph returned an existing event for a repeated transactionId", async () => {
    const at = (createdDateTime: string) => ({ ...createdReply, createdDateTime });
    // Made minutes before this request: an earlier identical create made it.
    expect((await create({}, at("2026-10-22T14:50:00Z"))).result.matchedExisting).toBe(true);
    // Made just now (or a little off by clock skew): this call created it.
    expect((await create({}, at("2026-10-22T15:00:01Z"))).result.matchedExisting).toBe(false);
    expect((await create({}, at("2026-10-22T14:59:00Z"))).result.matchedExisting).toBe(false);
    // No creation time: assume created.
    expect((await create({}, createdReply)).result.matchedExisting).toBe(false);
  });

  it("reports a recurrence's zone by its IANA name, as written", async () => {
    const recurrence = {
      pattern: { type: "weekly" as const, interval: 1, daysOfWeek: ["monday" as const] },
      range: { type: "numbered" as const, startDate: "2026-11-23", numberOfOccurrences: 4 },
    };
    const graphEcho = {
      ...createdReply,
      recurrence: {
        pattern: { type: "weekly", interval: 1, daysOfWeek: ["monday"] },
        range: {
          type: "numbered",
          recurrenceTimeZone: "Pacific Standard Time",
          numberOfOccurrences: 4,
        },
      },
    };
    const { result } = await create({ recurrence }, graphEcho);
    expect(result.recurrence).toMatchObject({
      range: { recurrenceTimeZone: tz, numberOfOccurrences: 4 },
    });
    // An IANA name Graph echoes is kept as is.
    const iana = {
      ...graphEcho,
      recurrence: { ...graphEcho.recurrence, range: { recurrenceTimeZone: "Europe/London" } },
    };
    expect((await create({ recurrence }, iana)).result.recurrence).toMatchObject({
      range: { recurrenceTimeZone: "Europe/London" },
    });
    // A matched series may have been written in another zone (an explicit
    // transactionId reused) or edited since: Graph's zone stands (Copilot review on #140).
    const matched = { ...graphEcho, createdDateTime: "2026-10-22T14:00:00Z" };
    const replay = await create(
      {
        start: { dateTime: "2026-11-23T09:00:00", timeZone: "Europe/London" },
        end: { dateTime: "2026-11-23T09:30:00", timeZone: "Europe/London" },
        recurrence,
      },
      matched,
    );
    expect(replay.result.matchedExisting).toBe(true);
    expect(replay.result.recurrence).toMatchObject({
      range: { recurrenceTimeZone: "Pacific Standard Time" },
    });
  });
});

describe("buildRecurrence (#51)", () => {
  const start = { dateTime: "2026-07-20T06:00:00", timeZone: "Australia/Brisbane" };
  const range = { type: "noEnd" as const, startDate: "2026-07-20" };
  const build = (pattern: RecurrenceInput["pattern"], r: Partial<RecurrenceInput["range"]> = {}) =>
    buildRecurrence({ pattern, range: { ...range, ...r } }, start);

  it("gives each pattern type exactly the fields it uses", () => {
    expect(
      build({ type: "daily", interval: 2, daysOfWeek: ["monday"], dayOfMonth: 3 }).pattern,
    ).toEqual({
      type: "daily",
      interval: 2,
    });
    expect(
      build({
        type: "weekly",
        interval: 1,
        daysOfWeek: ["friday"],
        firstDayOfWeek: "monday",
        month: 4,
      }).pattern,
    ).toEqual({ type: "weekly", interval: 1, daysOfWeek: ["friday"], firstDayOfWeek: "monday" });
    expect(build({ type: "absoluteMonthly", interval: 1, dayOfMonth: 15 }).pattern).toEqual({
      type: "absoluteMonthly",
      interval: 1,
      dayOfMonth: 15,
    });
    expect(
      build({ type: "relativeMonthly", interval: 1, daysOfWeek: ["monday"], index: "first" })
        .pattern,
    ).toEqual({ type: "relativeMonthly", interval: 1, daysOfWeek: ["monday"], index: "first" });
    expect(
      build({ type: "absoluteYearly", interval: 1, dayOfMonth: 25, month: 12 }).pattern,
    ).toEqual({
      type: "absoluteYearly",
      interval: 1,
      dayOfMonth: 25,
      month: 12,
    });
    expect(
      build({
        type: "relativeYearly",
        interval: 1,
        daysOfWeek: ["thursday"],
        index: "last",
        month: 11,
      }).pattern,
    ).toEqual({
      type: "relativeYearly",
      interval: 1,
      daysOfWeek: ["thursday"],
      index: "last",
      month: 11,
    });
  });

  it("refuses a pattern missing what its type needs", () => {
    expect(() => build({ type: "weekly", interval: 1 })).toThrow(
      /weekly recurrence needs pattern.daysOfWeek/,
    );
    expect(() => build({ type: "weekly", interval: 1, daysOfWeek: [] })).toThrow(/daysOfWeek/);
    expect(() => build({ type: "absoluteMonthly", interval: 1 })).toThrow(/pattern.dayOfMonth/);
    expect(() => build({ type: "relativeMonthly", interval: 1, daysOfWeek: ["monday"] })).toThrow(
      /pattern.index/,
    );
    expect(() => build({ type: "absoluteYearly", interval: 1, dayOfMonth: 1 })).toThrow(
      /pattern.month/,
    );
    expect(() =>
      build({ type: "relativeYearly", interval: 1, daysOfWeek: ["monday"], index: "first" }),
    ).toThrow(/pattern.month/);
    for (const interval of [0, 100, 1.5]) {
      expect(() => build({ type: "daily", interval })).toThrow(/interval/);
    }
  });

  it("starts the range on the event's start date, in its zone", () => {
    expect(build({ type: "daily", interval: 1 }).range).toEqual({
      type: "noEnd",
      startDate: "2026-07-20",
      recurrenceTimeZone: "Australia/Brisbane",
    });
    expect(() => build({ type: "daily", interval: 1 }, { startDate: "2026-07-21" })).toThrow(
      /must be the event's start date \(2026-07-20\)/,
    );
    for (const startDate of ["2026-7-20", "2026-02-30", "soon"]) {
      expect(() => build({ type: "daily", interval: 1 }, { startDate })).toThrow(ValidationError);
    }
  });

  it("checks the end of an endDate or numbered range", () => {
    const daily = { type: "daily" as const, interval: 1 };
    expect(build(daily, { type: "endDate", endDate: "2026-08-31" }).range).toMatchObject({
      endDate: "2026-08-31",
    });
    expect(build(daily, { type: "endDate", endDate: "2026-07-20" }).range).toMatchObject({
      endDate: "2026-07-20", // a single occurrence
    });
    expect(() => build(daily, { type: "endDate" })).toThrow(/needs range.endDate/);
    expect(() => build(daily, { type: "endDate", endDate: "2026-07-19" })).toThrow(
      /before its startDate/,
    );
    expect(() => build(daily, { type: "endDate", endDate: "2026-13-01" })).toThrow(ValidationError);
    expect(build(daily, { type: "numbered", numberOfOccurrences: 5 }).range).toMatchObject({
      numberOfOccurrences: 5,
    });
    for (const numberOfOccurrences of [undefined, 0, 1_000, 2.5]) {
      expect(() => build(daily, { type: "numbered", numberOfOccurrences })).toThrow(
        /numberOfOccurrences/,
      );
    }
    // A noEnd range carries neither.
    expect(build(daily, { endDate: "2026-08-31", numberOfOccurrences: 3 }).range).toEqual({
      type: "noEnd",
      startDate: "2026-07-20",
      recurrenceTimeZone: "Australia/Brisbane",
    });
  });
});

describe("end after start (#33)", () => {
  const brisbane = (dateTime: string) => ({ dateTime, timeZone: "Australia/Brisbane" });

  it("create_event refuses an end at or before the start, comparing instants", async () => {
    const post = vi.fn<PostMock>(async () => createdReply);
    for (const end of ["2026-07-20T09:00:00", "2026-07-20T10:00:00"]) {
      await expect(
        createEvent(graphWith({ post }), {
          subject: "x",
          start: brisbane("2026-07-20T10:00:00"),
          end: brisbane(end),
        }),
      ).rejects.toThrow(/end must be after its start/);
    }
    expect(post).not.toHaveBeenCalled();
    // An earlier wall-clock end in a zone further east can still be later.
    await createEvent(graphWith({ post }), {
      subject: "flight",
      start: brisbane("2026-07-20T10:00:00"), // 00:00Z
      end: { dateTime: "2026-07-20T09:30:00", timeZone: "Asia/Tokyo" }, // 00:30Z
    });
    expect(post).toHaveBeenCalledOnce();
  });

  it("update_event checks a new start against the event's current end, before writing", async () => {
    // Read in UTC (as Graph renders with Prefer UTC): 10:00-11:00 Brisbane.
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      "@odata.etag": 'W/"t1"',
      start: { dateTime: "2026-07-20T00:00:00.0000000", timeZone: "UTC" },
      end: { dateTime: "2026-07-20T01:00:00.0000000", timeZone: "UTC" },
    }));
    const patch = vi.fn<PatchMock>(async () => ({ id: "evt-1" }));
    await expect(
      updateEvent(graphWith({ get, patch }), { id: "evt-1", start: brisbane("2026-07-20T11:30") }),
    ).rejects.toThrow(/would end at or before it starts/);
    await expect(
      updateEvent(graphWith({ get, patch }), { id: "evt-1", end: brisbane("2026-07-20T09:00") }),
    ).rejects.toThrow(/would end at or before it starts/);
    expect(patch).not.toHaveBeenCalled();

    await updateEvent(graphWith({ get, patch }), {
      id: "evt-1",
      start: brisbane("2026-07-20T10:30"),
    });
    expect(patch).toHaveBeenCalledOnce();
  });

  it("knows which repeated daylight-saving hour a stored time is in (Copilot review on #116)", async () => {
    // New York repeats 01:00-02:00 on 2026-11-01. The stored end is the second
    // 01:30 (EST, 06:30Z); in local time it reads 01:30 either way.
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      "@odata.etag": 'W/"t1"',
      start: { dateTime: "2026-11-01T05:00:00.0000000", timeZone: "UTC" }, // 01:00 EDT
      end: { dateTime: "2026-11-01T06:30:00.0000000", timeZone: "UTC" }, // 01:30 EST
    }));
    const patch = vi.fn<PatchMock>(async () => ({ id: "evt-1" }));
    // A new 01:45 start (first occurrence, 05:45Z) is still before that end.
    const res = await updateEvent(graphWith({ get, patch }), {
      id: "evt-1",
      start: { dateTime: "2026-11-01T01:45", timeZone: "America/New_York" },
    });
    expect(patch).toHaveBeenCalledOnce();
    // The read asked for UTC; the report shows the stored start in the new zone.
    expect((get.mock.calls[0]![1] as ReqOpts).headers?.prefer).toBe('outlook.timezone="UTC"');
    expect(res.changes[0]).toMatchObject({
      field: "start",
      before: { dateTime: "2026-11-01T01:00:00.0000000", timeZone: "America/New_York" },
    });
  });
});

describe("updateEvent", () => {
  it("builds a partial patch with only provided fields", async () => {
    const patch = vi.fn<PatchMock>(async () => ({ id: "evt-1", subject: "New title" }));
    await updateEvent(graphWith({ patch }), { id: "evt-1", subject: "New title" });
    const [path, body] = patch.mock.calls[0]!;
    expect(path).toBe("/me/events/evt-1");
    expect(body).toEqual({ subject: "New title" });
  });

  it("maps a new start through the timezone-safe builder", async () => {
    const patch = vi.fn<PatchMock>(async () => ({ id: "evt-1" }));
    await updateEvent(graphWith({ patch }), {
      id: "evt-1",
      start: { dateTime: "2026-07-20T12:00:00", timeZone: "Australia/Brisbane" },
    });
    const body = patch.mock.calls[0]![1] as Record<string, unknown>;
    expect(body.start).toEqual({ dateTime: "2026-07-20T12:00:00", timeZone: "Australia/Brisbane" });
  });

  it("url-encodes the id and rejects an empty patch", async () => {
    const patch = vi.fn<PatchMock>(async () => ({ id: "x" }));
    await updateEvent(graphWith({ patch }), { id: "AA/BB", subject: "y" });
    expect(patch.mock.calls[0]![0]).toBe(`/me/events/${encodeURIComponent("AA/BB")}`);

    await expect(updateEvent(graphWith({ patch }), { id: "evt-1" })).rejects.toBeInstanceOf(
      ValidationError,
    );
  });

  it("reads the changed fields first and reports before/after values", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      "@odata.etag": 'W/"t1"',
      subject: "Old title",
      start: { dateTime: "2026-07-20T00:00:00.0000000", timeZone: "UTC" }, // 10:00 Brisbane
    }));
    const patch = vi.fn<PatchMock>(async () => ({
      id: "evt-1",
      subject: "New title",
      start: { dateTime: "2026-07-20T12:00:00.0000000", timeZone: "Australia/Brisbane" },
    }));
    const res = await updateEvent(graphWith({ get, patch }), {
      id: "evt-1",
      subject: "New title",
      start: { dateTime: "2026-07-20T12:00:00", timeZone: "Australia/Brisbane" },
    });

    const [getPath, getOpts] = get.mock.calls[0]! as [string, ReqOpts];
    expect(getPath).toBe("/me/events/evt-1");
    // end too: the new start is checked against the current end (#33).
    expect(getOpts.query.$select).toBe("id,attendees,subject,start,end");
    // The read is in UTC (exact); the write renders in the change's zone, and
    // the report converts the stored value to it, so before/after compare.
    expect(getOpts.headers?.prefer).toBe('outlook.timezone="UTC"');
    expect((patch.mock.calls[0]![2] as ReqOpts).headers).toEqual({
      prefer: 'outlook.timezone="Australia/Brisbane"',
      "if-match": 'W/"t1"',
    });

    expect(res.changes).toEqual([
      { field: "subject", before: "Old title", after: "New title" },
      {
        field: "start",
        before: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
        after: { dateTime: "2026-07-20T12:00:00.0000000", timeZone: "Australia/Brisbane" },
      },
    ]);
    expect(res.event.subject).toBe("New title");
  });

  it("sends no Prefer header when no time changes", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      "@odata.etag": 'W/"t1"',
      location: { displayName: "A" },
    }));
    const patch = vi.fn<PatchMock>(async () => ({ id: "evt-1", location: { displayName: "B" } }));
    const res = await updateEvent(graphWith({ get, patch }), { id: "evt-1", location: "B" });
    expect((get.mock.calls[0]![1] as ReqOpts).headers).toBeUndefined();
    expect(res.changes).toEqual([{ field: "location", before: "A", after: "B" }]);
  });

  it("reports bodies as plain text, truncating long ones", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      "@odata.etag": 'W/"t1"',
      body: { contentType: "html", content: "<p>Old <b>notes</b></p>" },
    }));
    const patch = vi.fn<PatchMock>(async () => ({
      id: "evt-1",
      body: { contentType: "text", content: "x".repeat(600) },
    }));
    const res = await updateEvent(graphWith({ get, patch }), {
      id: "evt-1",
      body: "x".repeat(600),
    });
    expect(res.changes[0]!.before).toBe("Old notes");
    expect(res.changes[0]!.after).toBe(`${"x".repeat(500)}…`);
  });

  it("does not PATCH when the event can't be read (e.g. a stale id)", async () => {
    const get = vi.fn<GetMock>(async () => {
      throw new Error("404");
    });
    const patch = vi.fn<PatchMock>();
    await expect(
      updateEvent(graphWith({ get, patch }), { id: "gone", subject: "x" }),
    ).rejects.toThrow("404");
    expect(patch).not.toHaveBeenCalled();
  });

  it("refuses events with attendees: a meeting update would email them (no PATCH)", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      subject: "Team sync",
      attendees: [{ emailAddress: { address: "colleague@example.com" } }],
    }));
    const patch = vi.fn<PatchMock>();
    await expect(
      updateEvent(graphWith({ get, patch }), { id: "evt-1", body: "copied mailbox text" }),
    ).rejects.toThrow(/has attendees/);
    expect((get.mock.calls[0]![1] as ReqOpts).query.$select).toContain("attendees");
    expect(patch).not.toHaveBeenCalled();
  });

  it("prepares the before-values before writing, so a conversion failure can't hide a write", async () => {
    // An invalid numeric entity makes htmlToText throw (until #32 lands).
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      body: { contentType: "html", content: "<p>&#x110000;</p>" },
    }));
    const patch = vi.fn<PatchMock>();
    await expect(
      updateEvent(graphWith({ get, patch }), { id: "evt-1", body: "x" }),
    ).rejects.toThrow();
    expect(patch).not.toHaveBeenCalled();
  });

  it("sends If-Match for the checked version and reports a 412 as 'changed meanwhile'", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
      "@odata.etag": 'W/"v1"',
      subject: "A",
    }));
    const ok = vi.fn<PatchMock>(async () => ({ id: "evt-1", subject: "B" }));
    await updateEvent(graphWith({ get, patch: ok }), { id: "evt-1", subject: "B" });
    expect((ok.mock.calls[0]![2] as ReqOpts).headers?.["if-match"]).toBe('W/"v1"');

    const conflict = vi.fn<PatchMock>(async () => {
      throw new GraphError("precondition failed", { status: 412 });
    });
    await expect(
      updateEvent(graphWith({ get, patch: conflict }), { id: "evt-1", subject: "B" }),
    ).rejects.toThrow(/no longer matches the version that was checked/);
  });
});

describe("cancelEvent (guarded soft delete, #55)", () => {
  type DeleteMock = (path: string, opts?: unknown) => Promise<void>;
  const trip = {
    id: "evt-1",
    "@odata.etag": 'W/"t1"',
    subject: "Travel to Bundaberg",
    start: { dateTime: "2026-07-20T09:00:00.0000000", timeZone: "Australia/Brisbane" },
    end: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
    type: "singleInstance",
    attendees: [],
  };
  const input = {
    id: "evt-1",
    confirmSubject: "travel to bundaberg ",
    timeZone: "Australia/Brisbane",
  };

  function graph(event: object) {
    const get = vi.fn<GetMock>(async () => event);
    const del = vi.fn<DeleteMock>(async () => undefined);
    return { get, del, client: { get, delete: del } as unknown as GraphClient };
  }

  it("checks the live event, then deletes it and returns it for the echo", async () => {
    const g = graph(trip);
    const removed = await cancelEvent(g.client, input);
    const [path, opts] = g.get.mock.calls[0]! as [string, ReqOpts];
    expect(path).toBe("/me/events/evt-1");
    expect(opts.query.$select).toContain("attendees");
    expect(opts.headers?.prefer).toBe('outlook.timezone="Australia/Brisbane"');
    expect(g.del.mock.calls[0]![0]).toBe("/me/events/evt-1");
    expect(removed).toMatchObject({ id: "evt-1", subject: "Travel to Bundaberg" });
  });

  it("addresses the event within another calendar when given its calendarId (#49)", async () => {
    const g = graph(trip);
    await cancelEvent(g.client, { ...input, calendarId: "cal/2" });
    expect(g.get.mock.calls[0]![0]).toBe("/me/calendars/cal%2F2/events/evt-1");
    expect(g.del.mock.calls[0]![0]).toBe("/me/calendars/cal%2F2/events/evt-1");
  });

  it("deletes only the checked version (If-Match) and reports a 412 as 'changed meanwhile'", async () => {
    const g = graph({ ...trip, "@odata.etag": 'W/"v7"' });
    await cancelEvent(g.client, input);
    expect((g.del.mock.calls[0]![1] as ReqOpts).headers?.["if-match"]).toBe('W/"v7"');

    const get = vi.fn<GetMock>(async () => ({ ...trip, "@odata.etag": 'W/"v7"' }));
    const del = vi.fn<DeleteMock>(async () => {
      throw new GraphError("precondition failed", { status: 412 });
    });
    await expect(
      cancelEvent({ get, delete: del } as unknown as GraphClient, input),
    ).rejects.toThrow(/an earlier attempt of this request already went through/);
  });

  it("refuses when confirmSubject doesn't match the title (nothing deleted)", async () => {
    const g = graph(trip);
    await expect(cancelEvent(g.client, { ...input, confirmSubject: "Dentist" })).rejects.toThrow(
      /does not match the event's title "Travel to Bundaberg"/,
    );
    expect(g.del).not.toHaveBeenCalled();
  });

  it("refuses events with attendees (deleting a meeting can email cancellations)", async () => {
    const g = graph({ ...trip, attendees: [{ emailAddress: { address: "a@example.com" } }] });
    await expect(cancelEvent(g.client, input)).rejects.toThrow(/has attendees/);
    expect(g.del).not.toHaveBeenCalled();
  });

  it("refuses anything recurring: a whole series, an occurrence, or an exception", async () => {
    const series = graph({ ...trip, type: "seriesMaster" });
    await expect(cancelEvent(series.client, input)).rejects.toThrow(/whole recurring series/);
    expect(series.del).not.toHaveBeenCalled();

    // Deleting an occurrence removes it from the series instead of using Deleted Items.
    for (const type of ["occurrence", "exception"]) {
      const g = graph({ ...trip, type });
      await expect(cancelEvent(g.client, input)).rejects.toThrow(/one occurrence/);
      expect(g.del).not.toHaveBeenCalled();
    }
  });

  it("rejects a non-IANA time zone before calling Graph", async () => {
    const g = graph(trip);
    await expect(cancelEvent(g.client, { ...input, timeZone: "Not/AZone" })).rejects.toThrow();
    expect(g.get).not.toHaveBeenCalled();
  });
});

describe("checked-version writes need a version tag", () => {
  it("refuses to PATCH or DELETE when Graph returns no @odata.etag", async () => {
    const untagged = { id: "evt-1", subject: "Trip", type: "singleInstance", attendees: [] };
    const get = vi.fn<GetMock>(async () => untagged);
    const patch = vi.fn<PatchMock>();
    const del = vi.fn();
    const client = { get, patch, delete: del } as unknown as GraphClient;

    await expect(updateEvent(client, { id: "evt-1", subject: "x" })).rejects.toThrow(
      /didn't return its version tag/,
    );
    await expect(
      cancelEvent(client, { id: "evt-1", confirmSubject: "Trip", timeZone: "Australia/Brisbane" }),
    ).rejects.toThrow(/didn't return its version tag/);
    expect(patch).not.toHaveBeenCalled();
    expect(del).not.toHaveBeenCalled();
  });

  it("doesn't claim 'nothing changed' on a 412: a retried request may already have applied", async () => {
    // e.g. the first PATCH committed but its response was lost; the retry then gets 412.
    const get = vi.fn<GetMock>(async () => ({ id: "evt-1", "@odata.etag": 'W/"t1"' }));
    const patch = vi.fn<PatchMock>(async () => {
      throw new GraphError("precondition failed", { status: 412 });
    });
    const err: Error = await updateEvent(graphWith({ get, patch }), {
      id: "evt-1",
      subject: "B",
    }).then(
      () => new Error("expected a rejection"),
      (e: unknown) => e as Error,
    );
    expect(err.message).toMatch(/already went through/);
    expect(err.message).not.toMatch(/nothing was changed/i);
  });
});

describe("no attendees / invitations (spec v2 §7.5, #61)", () => {
  const sneaky = {
    attendees: [{ emailAddress: { address: "someone@example.com" }, type: "required" }],
    isOnlineMeeting: true,
    responseRequested: true,
  };

  it("never allows attendee or meeting properties", () => {
    for (const f of ["attendees", "isOnlineMeeting", "responseRequested", "organizer"]) {
      expect(WRITABLE_EVENT_FIELDS.has(f)).toBe(false);
    }
    expect(() => assertWritableEventPayload({ subject: "x", attendees: [] })).toThrow(
      /Refusing to write event properties: attendees/,
    );
  });

  it("create_event drops attendee fields even if a caller smuggles them in", async () => {
    const post = vi.fn<PostMock>(async () => createdReply);
    await createEvent(graphWith({ post }), {
      subject: "Trip",
      start: { dateTime: "2026-07-20T10:00:00", timeZone: "Australia/Brisbane" },
      end: { dateTime: "2026-07-20T11:00:00", timeZone: "Australia/Brisbane" },
      ...sneaky,
    } as never);
    const payload = post.mock.calls[0]![1] as Record<string, unknown>;
    for (const k of Object.keys(payload)) expect(WRITABLE_EVENT_FIELDS.has(k)).toBe(true);
    expect(payload).not.toHaveProperty("attendees");
  });

  it("update_event drops attendee fields even if a caller smuggles them in", async () => {
    const patch = vi.fn<PatchMock>(async () => ({ id: "evt-1" }));
    await updateEvent(graphWith({ patch }), { id: "evt-1", subject: "x", ...sneaky } as never);
    expect(patch.mock.calls[0]![1]).toEqual({ subject: "x" });
  });
});
