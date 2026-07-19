import { describe, expect, it, vi } from "vitest";
import { z } from "zod";
import type { GraphClient } from "../graph/client.js";
import { getMessageTool, listMessagesTool, mailTools } from "./mail.js";
import type { ToolContext } from "./types.js";

function ctxWith(graph: Partial<GraphClient>): ToolContext {
  return { graph: graph as GraphClient, config: {} as ToolContext["config"] };
}

describe("mail tool schemas", () => {
  it("list_messages: defaults limit to 10 and caps at 25", () => {
    const schema = z.object(listMessagesTool.inputShape);
    expect(schema.parse({}).limit).toBe(10);
    expect(() => schema.parse({ limit: 26 })).toThrow();
    expect(() => schema.parse({ limit: 0 })).toThrow();
  });

  it("get_message: requires id and defaults format to text", () => {
    const schema = z.object(getMessageTool.inputShape);
    expect(() => schema.parse({})).toThrow();
    expect(schema.parse({ id: "abc" }).format).toBe("text");
    expect(() => schema.parse({ id: "abc", format: "pdf" })).toThrow();
  });

  it("exposes both tools as read-only", () => {
    expect(mailTools.map((t) => t.name).sort()).toEqual(["get_message", "list_messages"]);
    for (const t of mailTools) expect(t.annotations?.readOnlyHint).toBe(true);
  });
});

describe("list_messages handler", () => {
  it("summarises results and returns structured content", async () => {
    const request = vi.fn(async () => ({
      status: 200,
      data: {
        value: [{ id: "m1", subject: "Hi", hasAttachments: false }],
        "@odata.nextLink": "https://graph/next",
      },
      headers: new Headers(),
    }));
    const res = await listMessagesTool.handler(
      { query: "hi", limit: 10 },
      ctxWith({ request: request as unknown as GraphClient["request"] }),
    );
    expect(res.content[0]!.text).toContain("Found 1 message");
    expect(res.content[0]!.text).toContain("nextPageToken");
    expect(res.structuredContent).toMatchObject({ count: 1, nextPageToken: "https://graph/next" });
  });

  it("summarises an empty result", async () => {
    const request = vi.fn(async () => ({
      status: 200,
      data: { value: [] },
      headers: new Headers(),
    }));
    const res = await listMessagesTool.handler(
      { limit: 10 },
      ctxWith({ request: request as unknown as GraphClient["request"] }),
    );
    expect(res.content[0]!.text).toContain("No matching messages");
  });
});
