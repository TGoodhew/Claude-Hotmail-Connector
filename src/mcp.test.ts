import { existsSync } from "node:fs";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { describe, expect, it, vi } from "vitest";
import { loadConfig } from "./config.js";
import type { GraphClient } from "./graph/client.js";
import { allTools, createMcpServer } from "./mcp.js";
import type { ToolContext } from "./tools/types.js";

const EXPECTED_TOOLS = [
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
  it("advertises all six tools with input schemas", async () => {
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
    "lists all six tools over a real stdio handshake",
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
