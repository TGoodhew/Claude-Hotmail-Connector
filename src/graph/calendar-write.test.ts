import { describe, expect, it, vi } from "vitest";
import { ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { createEvent, updateEvent } from "./calendar.js";

type PostMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;
type PatchMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;
type GetMock = (path: string, opts?: unknown) => Promise<unknown>;
type ReqOpts = { query: Record<string, string>; headers?: Record<string, string> };

function graphWith(fns: { post?: PostMock; patch?: PatchMock; get?: GetMock }): GraphClient {
  // update_event reads the current values first; default to an empty event.
  return { get: async () => ({ id: "evt-1" }), ...fns } as unknown as GraphClient;
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
      subject: "Old title",
      start: { dateTime: "2026-07-20T10:00:00.0000000", timeZone: "Australia/Brisbane" },
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
    expect(getOpts.query.$select).toBe("id,subject,start");
    // Same zone on both requests, so before/after times are comparable.
    expect(getOpts.headers?.prefer).toBe('outlook.timezone="Australia/Brisbane"');
    expect((patch.mock.calls[0]![2] as ReqOpts).headers).toEqual(getOpts.headers);

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
    const get = vi.fn<GetMock>(async () => ({ id: "evt-1", location: { displayName: "A" } }));
    const patch = vi.fn<PatchMock>(async () => ({ id: "evt-1", location: { displayName: "B" } }));
    const res = await updateEvent(graphWith({ get, patch }), { id: "evt-1", location: "B" });
    expect((get.mock.calls[0]![1] as ReqOpts).headers).toBeUndefined();
    expect(res.changes).toEqual([{ field: "location", before: "A", after: "B" }]);
  });

  it("reports bodies as plain text, truncating long ones", async () => {
    const get = vi.fn<GetMock>(async () => ({
      id: "evt-1",
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
});
