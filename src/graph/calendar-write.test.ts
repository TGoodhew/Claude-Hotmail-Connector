import { describe, expect, it, vi } from "vitest";
import { ValidationError } from "../util/errors.js";
import type { GraphClient } from "./client.js";
import { createEvent, updateEvent } from "./calendar.js";

type PostMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;
type PatchMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;

function graphWith(fns: { post?: PostMock; patch?: PatchMock }): GraphClient {
  return fns as unknown as GraphClient;
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
});
