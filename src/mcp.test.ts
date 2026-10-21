import { existsSync } from "node:fs";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { describe, expect, it, vi } from "vitest";
import { loadConfig } from "./config.js";
import type { GraphClient } from "./graph/client.js";
import { allTools, createMcpServer, SERVER_INSTRUCTIONS, toolsForScopes } from "./mcp.js";
import type { ToolContext } from "./tools/types.js";

describe("toolsForScopes", () => {
  const names = (scopes: string[]) =>
    toolsForScopes(scopes)
      .map((t) => t.name)
      .sort();

  it("keeps every tool for the default scopes", () => {
    expect(toolsForScopes(loadConfig({ MICROSOFT_CLIENT_ID: "c" }).scopes)).toHaveLength(
      allTools().length,
    );
  });

  it("drops the calendar write tools when narrowed to Calendars.Read", () => {
    expect(names(["openid", "offline_access", "User.Read", "Mail.Read", "Calendars.Read"])).toEqual(
      ["get_message", "list_events", "list_messages", "whoami"],
    );
  });

  it("drops mail tools without Mail.Read, and whoami without User.Read", () => {
    expect(names(["openid", "offline_access", "Calendars.ReadWrite"])).toEqual([
      "cancel_event",
      "create_event",
      "list_events",
      "update_event",
    ]);
  });

  it("declares a scope for every tool", () => {
    for (const t of allTools()) expect(t.scopesAnyOf.length).toBeGreaterThan(0);
  });
});

const EXPECTED_TOOLS = [
  "cancel_event",
  "create_event",
  "get_message",
  "list_events",
  "list_messages",
  "update_event",
  "whoami",
];

const config = loadConfig({ MICROSOFT_CLIENT_ID: "test-client" });

function testCtx(): ToolContext {
  const graph = {
    get: vi.fn(async () => ({ displayName: "Tony", mail: "tony_goodhew@hotmail.com" })),
  } as unknown as GraphClient;
  return { graph, config };
}

describe("createMcpServer (in-memory transport)", () => {
  it("advertises all seven tools with input schemas", async () => {
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const server = createMcpServer(testCtx());
    await server.connect(serverTransport);

    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);

    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual(EXPECTED_TOOLS);
    for (const t of tools) {
      expect(t.description).toBeTruthy();
      expect(t.inputSchema).toBeTruthy();
    }

    await client.close();
    await server.close();
  });

  it("tells the client, at initialization, that mail and calendar content is untrusted (#88)", async () => {
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const server = createMcpServer(testCtx());
    await server.connect(serverTransport);
    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);

    expect(client.getInstructions()).toBe(SERVER_INSTRUCTIONS);
    expect(SERVER_INSTRUCTIONS).toMatch(/untrusted/);

    await client.close();
    await server.close();
  });

  it("executes a tool call end-to-end (whoami)", async () => {
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const server = createMcpServer(testCtx());
    await server.connect(serverTransport);
    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);

    const result = await client.callTool({ name: "whoami", arguments: {} });
    expect(JSON.stringify(result.content)).toContain("tony_goodhew@hotmail.com");

    await client.close();
    await server.close();
  });

  it("returns a redacted error result when a handler throws", async () => {
    const throwingGraph = {
      get: vi.fn(async () => {
        throw new Error("boom Bearer secret-token");
      }),
    } as unknown as GraphClient;
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const server = createMcpServer({ graph: throwingGraph, config });
    await server.connect(serverTransport);
    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);

    const result = await client.callTool({ name: "whoami", arguments: {} });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result.content)).not.toContain("secret-token");

    await client.close();
    await server.close();
  });
});

describe("allTools", () => {
  it("returns exactly the expected tool set", () => {
    expect(
      allTools()
        .map((t) => t.name)
        .sort(),
    ).toEqual(EXPECTED_TOOLS);
  });
});

// Real subprocess handshake — runs only when the server has been built.
const distPath = join(process.cwd(), "dist", "index.js");

function stringEnv(): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(process.env)) if (typeof v === "string") out[k] = v;
  return out;
}

describe("built server subprocess", () => {
  it.runIf(existsSync(distPath))(
    "lists all seven tools over a real stdio handshake",
    async () => {
      const { StdioClientTransport } = await import("@modelcontextprotocol/sdk/client/stdio.js");
      const transport = new StdioClientTransport({
        command: process.execPath,
        args: [distPath],
        env: { ...stringEnv(), MICROSOFT_CLIENT_ID: "smoke-test-dummy-id" },
        stderr: "ignore",
      });
      const client = new Client({ name: "smoke", version: "0.0.0" });
      await client.connect(transport);
      try {
        const { tools } = await client.listTools();
        expect(tools.map((t) => t.name).sort()).toEqual(EXPECTED_TOOLS);
      } finally {
        await client.close();
      }
    },
    30_000,
  );
});
