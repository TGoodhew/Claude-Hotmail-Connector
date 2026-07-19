import { describe, expect, it, vi } from "vitest";
import { z } from "zod";
import type { AppConfig } from "../config.js";
import type { GraphClient } from "../graph/client.js";
import { calendarTools, listEventsTool } from "./calendar.js";
import type { ToolContext } from "./types.js";

function ctxWith(graph: Partial<GraphClient>, defaultTimeZone = "Australia/Brisbane"): ToolContext {
  return { graph: graph as GraphClient, config: { defaultTimeZone } as AppConfig };
}

describe("list_events schema", () => {
  it("requires start and end, defaults limit, caps at 100", () => {
    const schema = z.object(listEventsTool.inputShape);
    expect(() => schema.parse({})).toThrow();
    expect(schema.parse({ startDateTime: "a", endDateTime: "b" }).limit).toBe(25);
    expect(() => schema.parse({ startDateTime: "a", endDateTime: "b", limit: 101 })).toThrow();
  });

  it("is registered read-only", () => {
    expect(calendarTools.map((t) => t.name)).toContain("list_events");
    expect(listEventsTool.annotations?.readOnlyHint).toBe(true);
  });
});

describe("list_events handler", () => {
  it("falls back to the config default time zone and summarises", async () => {
    const getAllPages = vi.fn(async () => [
      { id: "e1", subject: "Trip", start: {}, end: {}, isAllDay: false },
    ]);
    const res = await listEventsTool.handler(
      { startDateTime: "2026-07-20T00:00:00Z", endDateTime: "2026-07-21T00:00:00Z", limit: 25 },
      ctxWith(
        { getAllPages: getAllPages as unknown as GraphClient["getAllPages"] },
        "Europe/London",
      ),
    );
    expect(res.content[0]!.text).toContain("Europe/London");
    expect(res.structuredContent).toMatchObject({ count: 1, timeZone: "Europe/London" });
  });
});
