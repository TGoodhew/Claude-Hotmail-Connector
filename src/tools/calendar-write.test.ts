import { describe, expect, it, vi } from "vitest";
import { z } from "zod";
import type { AppConfig } from "../config.js";
import type { GraphClient } from "../graph/client.js";
import { calendarTools, createEventTool, updateEventTool } from "./calendar.js";
import type { ToolContext } from "./types.js";

function ctxWith(graph: Partial<GraphClient>, defaultTimeZone = "Australia/Brisbane"): ToolContext {
  return { graph: graph as GraphClient, config: { defaultTimeZone } as AppConfig };
}

type PostMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;
type PatchMock = (path: string, body: unknown, opts?: unknown) => Promise<unknown>;

describe("create_event tool", () => {
  it("registers three calendar tools and marks writes non-read-only + idempotent", () => {
    expect(calendarTools.map((t) => t.name).sort()).toEqual([
      "create_event",
      "list_events",
      "update_event",
    ]);
    expect(createEventTool.annotations?.readOnlyHint).toBe(false);
    expect(createEventTool.annotations?.idempotentHint).toBe(true);
  });

  it("description tells the model to confirm before calling", () => {
    expect(createEventTool.description.toLowerCase()).toContain("confirm");
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
});

describe("update_event tool", () => {
  it("passes through only provided fields", async () => {
    const patch = vi.fn<PatchMock>(async () => ({ id: "e1", subject: "Renamed" }));
    const res = await updateEventTool.handler(
      { id: "e1", subject: "Renamed" },
      ctxWith({ patch: patch as unknown as GraphClient["patch"] }),
    );
    expect(patch.mock.calls[0]![1]).toEqual({ subject: "Renamed" });
    expect(res.content[0]!.text).toContain("Updated event");
  });
});
