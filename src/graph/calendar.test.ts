import { describe, expect, it, vi } from "vitest";
import { ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { getEvent, listCalendars, listEvents, MAX_CALENDARS } from "./calendar.js";

describe("getEvent (#50)", () => {
  type GetMock = (
    path: string,
    opts?: { query?: Record<string, string>; headers?: Record<string, string> },
  ) => Promise<unknown>;
  const graphGet = (get: GetMock) => ({ get }) as unknown as GraphClient;
  const series = {
    id: "s1",
    subject: "Swim squad",
    start: { dateTime: "2026-07-20T06:00:00.0000000", timeZone: "Australia/Brisbane" },
    end: { dateTime: "2026-07-20T07:00:00.0000000", timeZone: "Australia/Brisbane" },
    type: "seriesMaster",
    showAs: "busy",
    organizer: { emailAddress: { name: "Coach", address: "coach@example.com" } },
    body: { contentType: "html", content: "<p>Bring <b>fins</b></p>" },
    attendees: [
      {
        type: "required",
        status: { response: "accepted" },
        emailAddress: { name: "Sam", address: "sam@example.com" },
      },
      {},
    ],
    isReminderOn: true,
    reminderMinutesBeforeStart: 30,
    recurrence: {
      pattern: { type: "weekly", interval: 1, daysOfWeek: ["monday"] },
      range: { type: "noEnd", startDate: "2026-07-20" },
    },
  };

  it("reads the full event in the requested zone, with its notes as text", async () => {
    const get = vi.fn<GetMock>(async () => series);
    const e = await getEvent(graphGet(get), { id: "s/1", timeZone: "Australia/Brisbane" });
    const [path, opts] = get.mock.calls[0]!;
    expect(path).toBe("/me/events/s%2F1");
    expect(opts?.headers?.prefer).toBe('outlook.timezone="Australia/Brisbane"');
    const select = String(opts?.query?.$select).split(",");
    for (const f of ["body", "attendees", "recurrence", "isReminderOn", "seriesMasterId"]) {
      expect(select).toContain(f);
    }
    expect(e).toMatchObject({
      id: "s1",
      subject: "Swim squad",
      type: "seriesMaster",
      showAs: "busy",
      organizer: { name: "Coach", address: "coach@example.com" },
      body: { contentType: "text", content: "Bring fins", truncated: false },
      attendees: [
        { name: "Sam", address: "sam@example.com", type: "required", response: "accepted" },
        { name: undefined, address: undefined, type: null, response: null },
      ],
      isReminderOn: true,
      reminderMinutesBeforeStart: 30,
      recurrence: series.recurrence,
    });
  });

  it("caps long notes and reads on from startChar, like get_message", async () => {
    const long = { id: "e1", body: { contentType: "text", content: "x".repeat(2_500) } };
    const get = vi.fn<GetMock>(async () => long);
    const first = await getEvent(graphGet(get), { id: "e1", timeZone: "UTC", maxChars: 1_000 });
    expect(first.body).toMatchObject({ truncated: true, totalChars: 2_500, nextStartChar: 1_000 });
    const next = await getEvent(graphGet(get), {
      id: "e1",
      timeZone: "UTC",
      maxChars: 1_000,
      startChar: 2_000,
    });
    expect(next.body).toMatchObject({ truncated: false, startChar: 2_000 });
    expect(next.body.content).toHaveLength(500);
  });

  it("gives empty values for an event without notes, attendees or a pattern", async () => {
    const e = await getEvent(
      graphGet(async () => ({ id: "e2" })),
      { id: "e2", timeZone: "UTC" },
    );
    expect(e).toMatchObject({
      body: { content: "", truncated: false },
      attendees: [],
      isReminderOn: null,
      reminderMinutesBeforeStart: null,
      recurrence: null,
    });
  });

  it("reads within another calendar, and rejects a bad zone before calling Graph", async () => {
    const get = vi.fn<GetMock>(async () => ({ id: "e3" }));
    await getEvent(graphGet(get), { id: "e3", timeZone: "UTC", calendarId: "cal/2" });
    expect(get.mock.calls[0]![0]).toBe("/me/calendars/cal%2F2/events/e3");
    await expect(
      getEvent(graphGet(get), { id: "e3", timeZone: "Nowhere/Nope" }),
    ).rejects.toBeInstanceOf(ValidationError);
    expect(get).toHaveBeenCalledTimes(1);
  });
});

type RequestMock = (
  method: string,
  path: string,
  opts?: { query?: Record<string, unknown>; headers?: Record<string, string> },
) => Promise<{ status: number; data: unknown; headers: Headers }>;

function graphWith(request: RequestMock): GraphClient {
  return { request } as unknown as GraphClient;
}

/** A calendarView page answer. */
function page(value: unknown[] = [], nextLink?: string) {
  return {
    status: 200,
    data: { value, ...(nextLink ? { "@odata.nextLink": nextLink } : {}) },
    headers: new Headers(),
  };
}

const validRange = {
  startDateTime: "2026-07-20T00:00:00+10:00",
  endDateTime: "2026-07-21T00:00:00+10:00",
  timeZone: "Australia/Brisbane",
  limit: 25,
};

const NEXT = "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=x&$skip=25";

describe("listEvents", () => {
  it("calls calendarView with the Prefer header and window params", async () => {
    const request = vi.fn<RequestMock>(async () => page());
    await listEvents(graphWith(request), validRange);

    const [method, path, opts] = request.mock.calls[0]!;
    expect(method).toBe("GET");
    expect(path).toBe("/me/calendarView");
    expect(opts?.headers?.prefer).toBe('outlook.timezone="Australia/Brisbane"');
    expect(opts?.query?.startDateTime).toBe(validRange.startDateTime);
    expect(opts?.query?.endDateTime).toBe(validRange.endDateTime);
    expect(opts?.query?.$orderby).toBe("start/dateTime");
    expect(opts?.query?.$top).toBe(25);
  });

  it("shapes events, with recurrence, free/busy, organizer and cancellation (#49)", async () => {
    const request = vi.fn<RequestMock>(async () =>
      page([
        {
          id: "e1",
          subject: "Macadamias Australia — visit",
          start: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
          end: { dateTime: "2026-07-20T11:00:00.0000000", timeZone: "Australia/Brisbane" },
          location: { displayName: "Bundaberg QLD" },
          isAllDay: false,
          webLink: "https://outlook/e1",
          type: "occurrence",
          seriesMasterId: "series-1",
          showAs: "tentative",
          organizer: { emailAddress: { name: "Tony", address: "tony@example.com" } },
          isCancelled: false,
        },
        { id: "e2" },
      ]),
    );
    const { events } = await listEvents(graphWith(request), validRange);
    expect(events).toEqual([
      {
        id: "e1",
        subject: "Macadamias Australia — visit",
        start: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
        end: { dateTime: "2026-07-20T11:00:00.0000000", timeZone: "Australia/Brisbane" },
        location: "Bundaberg QLD",
        isAllDay: false,
        webLink: "https://outlook/e1",
        type: "occurrence",
        seriesMasterId: "series-1",
        showAs: "tentative",
        organizer: { name: "Tony", address: "tony@example.com" },
        isCancelled: false,
      },
      {
        id: "e2",
        subject: null,
        start: null,
        end: null,
        location: null,
        isAllDay: false,
        webLink: null,
        type: null,
        seriesMasterId: null,
        showAs: null,
        organizer: null,
        isCancelled: false,
      },
    ]);
    const select = String(request.mock.calls[0]![2]?.query?.$select).split(",");
    for (const field of ["type", "seriesMasterId", "showAs", "organizer", "isCancelled"]) {
      expect(select).toContain(field);
    }
  });

  it("says when the range has more events than were returned, never silently capping (#49)", async () => {
    const request = vi.fn<RequestMock>(async () => page([{ id: "e1" }], NEXT));
    const res = await listEvents(graphWith(request), validRange);
    expect(res.truncated).toBe(true);
    expect(res.nextPageToken).toMatch(/^[A-Za-z0-9_-]+$/); // opaque, not a URL
    expect(res.nextPageToken).not.toContain("https://");

    const last = await listEvents(graphWith(vi.fn<RequestMock>(async () => page())), validRange);
    expect(last.truncated).toBe(false);
    expect(last.nextPageToken).toBeUndefined();
  });

  it("follows its cursor relative to Graph, keeping the display zone (#49)", async () => {
    const request = vi
      .fn<RequestMock>()
      .mockResolvedValueOnce(page([{ id: "e1" }], NEXT))
      .mockResolvedValueOnce(page([{ id: "e2" }]));
    const first = await listEvents(graphWith(request), {
      ...validRange,
      timeZone: "Europe/London",
    });
    // The other inputs don't matter with a pageToken: the cursor decides.
    const second = await listEvents(graphWith(request), {
      ...validRange,
      pageToken: first.nextPageToken,
    });
    const [, path, opts] = request.mock.calls[1]!;
    expect(path).toBe("me/calendarView?startDateTime=x&$skip=25");
    expect(opts?.query).toBeUndefined();
    expect(opts?.headers?.prefer).toBe('outlook.timezone="Europe/London"');
    expect(second).toMatchObject({ timeZone: "Europe/London", truncated: false });
    expect(second.events.map((e) => e.id)).toEqual(["e2"]);
  });

  it("is still truncated when Graph's nextLink can't be followed, but gives no cursor", async () => {
    const request = vi.fn<RequestMock>(async () => page([], "https://evil.example/next"));
    const res = await listEvents(graphWith(request), validRange);
    expect(res.truncated).toBe(true);
    expect(res.nextPageToken).toBeUndefined();
  });

  it("rejects a URL, a forged cursor, or a list_messages cursor as pageToken", async () => {
    const request = vi.fn<RequestMock>();
    const token = (o: unknown) => Buffer.from(JSON.stringify(o)).toString("base64url");
    for (const pageToken of [
      "https://evil.example/steal",
      token({ v: 1, next: "https://evil.example/x", kind: "events", timeZone: "UTC" }),
      token({ v: 1, next: "me/messages?$skip=10", mode: "filter" }),
    ]) {
      await expect(listEvents(graphWith(request), { ...validRange, pageToken })).rejects.toThrow(
        /not a valid cursor.*list_events/,
      );
    }
    await expect(
      listEvents(graphWith(request), {
        ...validRange,
        pageToken: token({ v: 1, next: "me/calendarView", kind: "events", timeZone: "Nope/Nope" }),
      }),
    ).rejects.toBeInstanceOf(ValidationError);
    expect(request).not.toHaveBeenCalled();
  });

  it("pins a bare window that spans Sydney's spring-forward to the right instants (#64)", async () => {
    // From @masyarakatai's #76: 00:00 (+10) to 05:00 (+11) on 2026-10-04 is
    // four real hours, not five.
    const request = vi.fn<RequestMock>(async () => page());
    await listEvents(graphWith(request), {
      startDateTime: "2026-10-04T00:00:00",
      endDateTime: "2026-10-04T05:00:00",
      timeZone: "Australia/Sydney",
      limit: 25,
    });
    const opts = request.mock.calls[0]![2];
    expect(opts?.query).toMatchObject({
      startDateTime: "2026-10-03T14:00:00.000Z",
      endDateTime: "2026-10-03T18:00:00.000Z",
    });
    expect(opts?.headers?.prefer).toBe('outlook.timezone="Australia/Sydney"');
  });

  it("lists another calendar by its id (#49)", async () => {
    const request = vi.fn<RequestMock>(async () => page());
    await listEvents(graphWith(request), { ...validRange, calendarId: "AAA/b+c=" });
    expect(request.mock.calls[0]![1]).toBe("/me/calendars/AAA%2Fb%2Bc%3D/calendarView");
  });

  it("rejects an invalid time zone", async () => {
    const request = vi.fn<RequestMock>(async () => page());
    await expect(
      listEvents(graphWith(request), { ...validRange, timeZone: "Nowhere/Nope" }),
    ).rejects.toBeInstanceOf(ValidationError);
  });

  it("rejects invalid range dates", async () => {
    const request = vi.fn<RequestMock>(async () => page());
    await expect(
      listEvents(graphWith(request), { ...validRange, startDateTime: "nope" }),
    ).rejects.toBeInstanceOf(ValidationError);
  });

  it("pins a date-only window to the zone's midnights (#33)", async () => {
    const request = vi.fn<RequestMock>(async () => page());
    await listEvents(graphWith(request), {
      ...validRange,
      startDateTime: "2026-07-20",
      endDateTime: "2026-07-21",
    });
    expect(request.mock.calls[0]![2]?.query).toMatchObject({
      startDateTime: "2026-07-19T14:00:00.000Z",
      endDateTime: "2026-07-20T14:00:00.000Z",
    });
  });

  it("rejects a window that ends at or before it starts, without calling Graph (#33)", async () => {
    const request = vi.fn<RequestMock>(async () => page());
    for (const endDateTime of ["2026-07-19T23:00:00+10:00", validRange.startDateTime]) {
      await expect(listEvents(graphWith(request), { ...validRange, endDateTime })).rejects.toThrow(
        /endDateTime must be after startDateTime/,
      );
    }
    expect(request).not.toHaveBeenCalled();
  });

  it("orders sub-second bounds correctly, local or mixed with an offset (Copilot review on #116)", async () => {
    const request = vi.fn<RequestMock>(async () => page());
    const run = (startDateTime: string, endDateTime: string) =>
      listEvents(graphWith(request), { ...validRange, startDateTime, endDateTime });

    // An 800 ms local window is valid, and both bounds keep their milliseconds.
    await run("2026-07-20T10:00:00.100", "2026-07-20T10:00:00.900");
    expect(request.mock.calls[0]![2]?.query).toMatchObject({
      startDateTime: "2026-07-20T00:00:00.100Z",
      endDateTime: "2026-07-20T00:00:00.900Z",
    });
    // A local start after an offset end, by 800 ms, is reversed.
    await expect(run("2026-07-20T10:00:00.900", "2026-07-20T10:00:00.100+10:00")).rejects.toThrow(
      /endDateTime must be after startDateTime/,
    );
    // And the mixed case the right way round is fine.
    await run("2026-07-20T10:00:00.100", "2026-07-20T10:00:00.900+10:00");
    expect(request).toHaveBeenCalledTimes(2);
  });
});

describe("listCalendars (#49)", () => {
  it("lists the user's calendars, own and shared", async () => {
    const getAllPages = vi.fn(async () => [
      {
        id: "c1",
        name: "Calendar",
        isDefaultCalendar: true,
        canEdit: true,
        owner: { name: "Tony", address: "tony@example.com" },
      },
      { id: "c2" },
    ]);
    const calendars = await listCalendars({ getAllPages } as unknown as GraphClient);
    expect(calendars).toEqual([
      {
        id: "c1",
        name: "Calendar",
        isDefaultCalendar: true,
        canEdit: true,
        owner: { name: "Tony", address: "tony@example.com" },
      },
      { id: "c2", name: null, isDefaultCalendar: false, canEdit: false, owner: null },
    ]);
    expect(getAllPages).toHaveBeenCalledWith(
      "/me/calendars",
      { query: { $select: "id,name,isDefaultCalendar,canEdit,owner", $top: MAX_CALENDARS } },
      { maxItems: MAX_CALENDARS },
    );
  });
});
