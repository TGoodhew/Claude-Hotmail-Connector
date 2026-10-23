import { describe, expect, it, vi } from "vitest";
import { z } from "zod";
import type { AppConfig } from "../config.js";
import type { GraphClient } from "../graph/client.js";
import { calendarTools, cancelEventTool, createEventTool, updateEventTool } from "./calendar.js";
import type { ToolContext } from "./types.js";

function ctxWith(graph: Partial<GraphClient>, defaultTimeZone = "Australia/Brisbane"): ToolContext {
  return { graph: graph as GraphClient, config: { defaultTimeZone } as AppConfig };
}

type PostMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;
type PatchMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;

describe("create_event tool", () => {
  it("registers seven calendar tools and marks writes non-read-only + idempotent", () => {
    expect(calendarTools.map((t) => t.name).sort()).toEqual([
      "cancel_event",
      "create_event",
      "find_free_time",
      "get_event",
      "list_calendars",
      "list_events",
      "update_event",
    ]);
    expect(createEventTool.annotations?.readOnlyHint).toBe(false);
    expect(createEventTool.annotations?.idempotentHint).toBe(true);
  });

  it("description tells the model to confirm before calling", () => {
    expect(createEventTool.description.toLowerCase()).toContain("confirm");
    // And that a mistaken series can't be undone with cancel_event (#51).
    expect(createEventTool.description).toContain("can't be removed with cancel_event");
  });

  it("passes an all-day, free, recurring event through (#51)", async () => {
    const post = vi.fn<PostMock>(async () => ({ id: "e1", subject: "Leave" }));
    const recurrence = {
      pattern: { type: "absoluteYearly" as const, interval: 1, dayOfMonth: 20, month: 7 },
      range: { type: "numbered" as const, startDate: "2026-07-20", numberOfOccurrences: 3 },
    };
    const res = await createEventTool.handler(
      {
        subject: "Leave",
        start: { dateTime: "2026-07-20" },
        end: { dateTime: "2026-07-21" },
        isAllDay: true,
        showAs: "oof",
        recurrence,
      },
      ctxWith({ post: post as unknown as GraphClient["post"] }, "Europe/London"),
    );
    expect(post.mock.calls[0]![1]).toMatchObject({
      isAllDay: true,
      showAs: "oof",
      start: { dateTime: "2026-07-20T00:00:00", timeZone: "Europe/London" },
      recurrence: { pattern: recurrence.pattern, range: { recurrenceTimeZone: "Europe/London" } },
    });
    expect(res.structuredContent).toMatchObject({ event: { id: "e1", isAllDay: false } });
  });

  it("says so when an identical create matched an existing event, rather than 'Created' (#135)", async () => {
    const created = (createdDateTime: string) =>
      vi.fn<PostMock>(async () => ({ id: "e1", subject: "Swim", createdDateTime }));
    const call = (post: ReturnType<typeof created>) =>
      createEventTool.handler(
        {
          subject: "Swim",
          start: { dateTime: "2026-11-23T09:00:00" },
          end: { dateTime: "2026-11-23T09:30:00" },
        },
        ctxWith({ post: post as unknown as GraphClient["post"] }),
      );
    const old = await call(created("2000-01-01T00:00:00Z"));
    expect(old.content[0]!.text).toMatch(/already exists \("Swim"\): nothing new was created/);
    expect(old.structuredContent).toMatchObject({ event: { matchedExisting: true } });
    const fresh = await call(created(new Date().toISOString()));
    expect(fresh.content[0]!.text).toMatch(/^Created event "Swim"\./);
  });

  it("schema rejects an unknown showAs or recurrence type", () => {
    const schema = z.object(createEventTool.inputShape);
    const base = { subject: "x", start: { dateTime: "a" }, end: { dateTime: "b" } };
    expect(() => schema.parse({ ...base, showAs: "invisible" })).toThrow();
    expect(() =>
      schema.parse({
        ...base,
        recurrence: {
          pattern: { type: "hourly", interval: 1 },
          range: { type: "noEnd", startDate: "x" },
        },
      }),
    ).toThrow();
  });

  it("defaults each time's zone to the config default", async () => {
    const post = vi.fn<PostMock>(async () => ({ id: "e1", subject: "Trip" }));
    await createEventTool.handler(
      {
        subject: "Trip",
        start: { dateTime: "2026-07-20T10:00:00" },
        end: { dateTime: "2026-07-20T11:00:00" },
      },
      ctxWith({ post: post as unknown as GraphClient["post"] }, "Europe/London"),
    );
    const payload = post.mock.calls[0]![1] as Record<string, unknown>;
    expect(payload.start).toEqual({ dateTime: "2026-07-20T10:00:00", timeZone: "Europe/London" });
  });

  it("schema requires subject/start/end", () => {
    const schema = z.object(createEventTool.inputShape);
    expect(() => schema.parse({})).toThrow();
    expect(() =>
      schema.parse({
        subject: "x",
        start: { dateTime: "a" },
        end: { dateTime: "b" },
      }),
    ).not.toThrow();
  });

  it("strips attendees from the arguments an MCP client sends (no invitations)", async () => {
    const post = vi.fn<PostMock>(async () => ({ id: "e1", subject: "Trip" }));
    const registered = calendarTools.find((t) => t.name === "create_event")!;
    await registered.handler(
      {
        subject: "Trip",
        start: { dateTime: "2026-07-20T10:00:00" },
        end: { dateTime: "2026-07-20T11:00:00" },
        attendees: [{ emailAddress: { address: "someone@example.com" } }],
      },
      ctxWith({ post: post as unknown as GraphClient["post"] }),
    );
    expect(post.mock.calls[0]![1]).not.toHaveProperty("attendees");
  });
});

describe("update_event tool", () => {
  it("passes through only provided fields and returns the before/after changes", async () => {
    const get = vi.fn(async () => ({ id: "e1", "@odata.etag": 'W/"t1"', subject: "Original" }));
    const patch = vi.fn<PatchMock>(async () => ({ id: "e1", subject: "Renamed" }));
    const res = await updateEventTool.handler(
      { id: "e1", subject: "Renamed" },
      ctxWith({
        get: get as unknown as GraphClient["get"],
        patch: patch as unknown as GraphClient["patch"],
      }),
    );
    expect(patch.mock.calls[0]![1]).toEqual({ subject: "Renamed" });
    expect(res.content[0]!.text).toContain("Updated event");
    expect(res.content[0]!.text).toContain("changed: subject");
    expect(res.structuredContent?.changes).toEqual([
      { field: "subject", before: "Original", after: "Renamed" },
    ]);
  });

  it("is marked destructive and tells the model to confirm and show the changes", () => {
    expect(updateEventTool.annotations?.destructiveHint).toBe(true);
    const d = updateEventTool.description.toLowerCase();
    expect(d).toContain("confirm");
    expect(d).toContain("before and after");
  });
});

describe("cancel_event tool", () => {
  it("is destructive, needs Calendars.ReadWrite, and demands explicit confirmation", () => {
    expect(cancelEventTool.annotations).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
    });
    expect(cancelEventTool.scopesAnyOf).toEqual(["Calendars.ReadWrite"]);
    const d = cancelEventTool.description.toLowerCase();
    expect(d).toContain("destructive");
    expect(d).toContain("explicitly confirmed");
    expect(z.object(cancelEventTool.inputShape).safeParse({ id: "e1" }).success).toBe(false);
  });

  it("deletes after the checks and reports the removed event", async () => {
    const get = vi.fn(async () => ({
      id: "e1",
      "@odata.etag": 'W/"t1"',
      subject: "Trip",
      attendees: [],
    }));
    const del = vi.fn(async (_path: string, _opts?: unknown) => undefined);
    const res = await cancelEventTool.handler(
      { id: "e1", confirmSubject: "Trip" },
      ctxWith({
        get: get as unknown as GraphClient["get"],
        delete: del as unknown as GraphClient["delete"],
      }),
    );
    expect(del.mock.calls[0]![0]).toBe("/me/events/e1");
    expect(res.content[0]!.text).toContain('Cancelled event "Trip"');
    expect(res.structuredContent?.cancelled).toMatchObject({ id: "e1", subject: "Trip" });
  });
});
