import { describe, expect, it, vi } from "vitest";
import { GraphError, ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import {
  assertWritableEventPayload,
  cancelEvent,
  createEvent,
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
