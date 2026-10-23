import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { describe, expect, it, vi } from "vitest";
import { loadConfig } from "./config.js";
import type { GraphClient } from "./graph/client.js";
import {
  allTools,
  createMcpServer,
  SERVER_INSTRUCTIONS,
  SERVER_VERSION,
  TOOL_CALL_DEADLINE_MS,
  toolsForScopes,
} from "./mcp.js";
import type { RegisteredTool, ToolContext } from "./tools/types.js";

describe("toolsForScopes", () => {
  const names = (scopes: string[]) =>
    toolsForScopes(scopes)
      .map((t) => t.name)
      .sort();

  it("keeps every tool but the send tools for the default scopes", () => {
    const send = ["reply_to_email", "send_email"];
    expect(names(loadConfig({ MICROSOFT_CLIENT_ID: "c" }).scopes)).toEqual(
      allTools()
        .map((t) => t.name)
        .filter((n) => !send.includes(n))
        .sort(),
    );
  });

  it("offers reply_to_email only with Mail.Read as well as Mail.Send (#57)", () => {
    const sendOnly = loadConfig({
      MICROSOFT_CLIENT_ID: "c",
      ALLOW_SEND_MAIL: "true",
      MICROSOFT_SCOPES: "Mail.Send",
    }).scopes;
    expect(names(sendOnly)).toContain("send_email");
    expect(names(sendOnly)).not.toContain("reply_to_email");
  });

  it("adds the send tools only when sending is turned on (#57)", () => {
    const on = loadConfig({ MICROSOFT_CLIENT_ID: "c", ALLOW_SEND_MAIL: "true" }).scopes;
    expect(toolsForScopes(on)).toHaveLength(allTools().length);
    // Narrowed away again, they go too.
    const narrowed = loadConfig({
      MICROSOFT_CLIENT_ID: "c",
      ALLOW_SEND_MAIL: "true",
      MICROSOFT_SCOPES: "User.Read Mail.Read",
    }).scopes;
    expect(names(narrowed)).not.toContain("send_email");
  });

  it("drops the calendar write tools when narrowed to Calendars.Read", () => {
    expect(names(["openid", "offline_access", "User.Read", "Mail.Read", "Calendars.Read"])).toEqual(
      [
        "find_free_time",
        "get_attachment",
        "get_event",
        "get_message",
        "list_calendars",
        "list_events",
        "list_mail_folders",
        "list_messages",
        "sign_out",
        "whoami",
      ],
    );
  });

  it("drops mail tools without Mail.Read, and whoami without User.Read", () => {
    expect(names(["openid", "offline_access", "Calendars.ReadWrite"])).toEqual([
      "cancel_event",
      "create_event",
      "find_free_time",
      "get_event",
      "list_calendars",
      "list_events",
      "sign_out",
      "update_event",
    ]);
  });

  it("says in every description that it is the personal Outlook.com / Hotmail account (#65)", () => {
    for (const t of allTools()) {
      expect(t.description, t.name).toContain("personal Outlook.com / Hotmail");
    }
  });

  it("declares a scope for every tool", () => {
    for (const t of allTools()) expect(t.scopesAnyOf.length).toBeGreaterThan(0);
  });
});

const EXPECTED_TOOLS = [
  "cancel_event",
  "create_event",
  "find_free_time",
  "get_attachment",
  "get_event",
  "get_message",
  "list_calendars",
  "list_events",
  "list_mail_folders",
  "list_messages",
  "reply_to_email",
  "send_email",
  "sign_out",
  "update_event",
  "whoami",
];

/** Offered by default: all but the send tools, which need ALLOW_SEND_MAIL (#57). */
const DEFAULT_TOOLS = EXPECTED_TOOLS.filter((n) => n !== "send_email" && n !== "reply_to_email");

const config = loadConfig({ MICROSOFT_CLIENT_ID: "test-client" });

function testCtx(): ToolContext {
  const graph = {
    get: vi.fn(async () => ({ displayName: "Tony", mail: "tony_goodhew@hotmail.com" })),
  } as unknown as GraphClient;
  return { graph, config };
}

describe("tool calls are bound to one sign-out epoch", () => {
  it("every Graph request of a call carries the epoch from the call's start", async () => {
    let epoch = 3;
    const seen: unknown[] = [];
    const graph = {
      get: vi.fn(async (_path: string, opts?: { signOutEpoch?: number }) => {
        seen.push(opts?.signOutEpoch);
        epoch = 4; // a sign-out happens between this call's requests
        return {};
      }),
    } as unknown as GraphClient;
    const twoRequests: RegisteredTool = {
      name: "two_requests",
      description: "test tool",
      inputShape: {},
      scopesAnyOf: ["openid"],
      handler: async (_input, ctx) => {
        await ctx.graph.get("/a");
        await ctx.graph.get("/b");
        return { content: [{ type: "text", text: "ok" }] };
      },
    };
    const server = createMcpServer(
      { graph, config, auth: { signOut: async () => undefined, signOutEpoch: () => epoch } },
      [twoRequests],
    );
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);
    await client.callTool({ name: "two_requests", arguments: {} });
    expect(seen).toEqual([3, 3]); // not 3 then 4
    await client.close();
  });

  it("catches up with sign-outs in other processes before noting the epoch (#95)", async () => {
    let epoch = 3;
    const seen: unknown[] = [];
    const graph = {
      get: vi.fn(async (_path: string, opts?: { signOutEpoch?: number }) => {
        seen.push(opts?.signOutEpoch);
        return {};
      }),
    } as unknown as GraphClient;
    const oneRequest: RegisteredTool = {
      name: "one_request",
      description: "test tool",
      inputShape: {},
      scopesAnyOf: ["openid"],
      handler: async (_input, ctx) => {
        await ctx.graph.get("/a");
        return { content: [{ type: "text", text: "ok" }] };
      },
    };
    const auth = {
      signOut: async () => undefined,
      signOutEpoch: () => epoch,
      // A sign-out elsewhere, first seen now: it comes before this call.
      syncSignOuts: vi.fn(async () => (epoch = 4)),
    };
    const server = createMcpServer({ graph, config, auth }, [oneRequest]);
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);
    await client.callTool({ name: "one_request", arguments: {} });
    expect(auth.syncSignOuts).toHaveBeenCalledTimes(1);
    expect(seen).toEqual([4]);
    await client.close();
  });

  it("stops waiting for a stalled catch-up when the client cancels the call (Copilot review on #126)", async () => {
    const handler = vi.fn(async () => ({ content: [{ type: "text" as const, text: "ok" }] }));
    const tool: RegisteredTool = {
      name: "stalled",
      description: "test tool",
      inputShape: {},
      scopesAnyOf: ["openid"],
      handler,
    };
    const auth = {
      signOut: async () => undefined,
      signOutEpoch: () => 0,
      syncSignOuts: vi.fn(() => new Promise<number>(() => undefined)), // never settles
    };
    const failures: string[] = [];
    const stderr = vi.spyOn(process.stderr, "write").mockImplementation((chunk) => {
      if (String(chunk).includes("Tool handler failed")) failures.push(String(chunk));
      return true;
    });
    try {
      const server = createMcpServer({ graph: {} as GraphClient, config, auth }, [tool]);
      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
      await server.connect(serverTransport);
      const client = new Client({ name: "test", version: "0.0.0" });
      await client.connect(clientTransport);
      const controller = new AbortController();
      const call = client
        .callTool({ name: "stalled", arguments: {} }, undefined, { signal: controller.signal })
        .catch(() => undefined);
      await vi.waitFor(() => expect(auth.syncSignOuts).toHaveBeenCalled());
      controller.abort();
      await call;
      // The server side gave up on the catch-up and ended the call.
      await vi.waitFor(() => expect(failures.join("")).toMatch(/cancelled/));
      expect(handler).not.toHaveBeenCalled();
      await client.close();
    } finally {
      stderr.mockRestore();
    }
  });

  it("gives every request the client's cancellation signal and one deadline (#37)", async () => {
    const seen: { signal?: AbortSignal; deadline?: number }[] = [];
    const graph = {
      get: vi.fn(async (_path: string, opts?: { signal?: AbortSignal; deadline?: number }) => {
        seen.push({ signal: opts?.signal, deadline: opts?.deadline });
        return {};
      }),
    } as unknown as GraphClient;
    const twoRequests: RegisteredTool = {
      name: "two_requests",
      description: "test tool",
      inputShape: {},
      scopesAnyOf: ["openid"],
      handler: async (_input, ctx) => {
        await ctx.graph.get("/a");
        await ctx.graph.get("/b");
        return { content: [{ type: "text", text: "ok" }] };
      },
    };
    const server = createMcpServer({ graph, config }, [twoRequests]);
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);
    const before = Date.now();
    await client.callTool({ name: "two_requests", arguments: {} });

    expect(seen).toHaveLength(2);
    expect(seen[0]!.signal).toBeInstanceOf(AbortSignal);
    expect(seen[1]!.signal).toBe(seen[0]!.signal);
    expect(seen[1]!.deadline).toBe(seen[0]!.deadline); // one deadline for the call
    expect(seen[0]!.deadline! - before).toBeLessThanOrEqual(TOOL_CALL_DEADLINE_MS + 1_000);
    expect(seen[0]!.deadline! - before).toBeGreaterThan(TOOL_CALL_DEADLINE_MS - 5_000);
    await client.close();
  });
});

describe("createMcpServer (in-memory transport)", () => {
  it("advertises the thirteen default tools with input schemas", async () => {
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const server = createMcpServer(testCtx());
    await server.connect(serverTransport);

    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);

    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual(DEFAULT_TOOLS);
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

describe("SERVER_VERSION (#46)", () => {
  it("matches package.json, which the release tag and .mcpb manifest also use", () => {
    const pkg = JSON.parse(readFileSync(join(process.cwd(), "package.json"), "utf8")) as {
      version: string;
    };
    expect(SERVER_VERSION).toBe(pkg.version);
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
    "lists the thirteen default tools over a real stdio handshake",
    async () => {
      const { StdioClientTransport } = await import("@modelcontextprotocol/sdk/client/stdio.js");
      const transport = new StdioClientTransport({
        command: process.execPath,
        args: [distPath],
        // Switches pinned off: a checkout's .env may turn them on (the environment wins).
        env: {
          ...stringEnv(),
          MICROSOFT_CLIENT_ID: "smoke-test-dummy-id",
          ALLOW_SEND_MAIL: "false",
          LOG_FILE: "false",
        },
        stderr: "ignore",
      });
      const client = new Client({ name: "smoke", version: "0.0.0" });
      await client.connect(transport);
      try {
        const { tools } = await client.listTools();
        expect(tools.map((t) => t.name).sort()).toEqual(DEFAULT_TOOLS);
      } finally {
        await client.close();
      }
    },
    30_000,
  );

  it.runIf(existsSync(distPath))(
    "writes a configuration error to the log file (#178)",
    () => {
      const dir = mkdtempSync(join(tmpdir(), "cht-startup-log-"));
      try {
        const run = spawnSync(process.execPath, [distPath], {
          env: {
            ...stringEnv(),
            MICROSOFT_CLIENT_ID: "smoke-test-dummy-id",
            DEFAULT_TIMEZONE: "Not/AZone",
            LOG_FILE: "true",
            CLAUDE_HOTMAIL_CACHE_DIR: dir,
          },
          encoding: "utf8",
          timeout: 20_000,
        });
        expect(run.status).toBe(1);
        const log = readFileSync(join(dir, "connector.log"), "utf8");
        expect(log).toContain('"msg":"Starting."');
        expect(log).toContain('"msg":"Stopped."');
        expect(log).toContain("Not/AZone");
      } finally {
        rmSync(dir, { recursive: true, force: true });
      }
    },
    30_000,
  );
});

describe("the client for work that outlives a call (#57)", () => {
  it("keeps the call's sign-out epoch, but not its deadline or cancellation", async () => {
    const seen: Record<string, unknown>[] = [];
    const graph = {
      get: vi.fn(async (_p: string, opts?: Record<string, unknown>) => {
        seen.push(opts ?? {});
        return {};
      }),
    } as unknown as GraphClient;
    let late: GraphClient | undefined;
    let deadline: number | undefined;
    const probe: RegisteredTool = {
      name: "probe",
      description: "test",
      inputShape: {},
      scopesAnyOf: ["User.Read"],
      async handler(_input, ctx) {
        late = ctx.unscopedGraph;
        deadline = ctx.deadline;
        return { content: [{ type: "text", text: "ok" }] };
      },
    };
    const ctx: ToolContext = {
      graph,
      config,
      auth: { signOut: async () => undefined, syncSignOuts: async () => 7 },
    };
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await createMcpServer(ctx, [probe]).connect(serverTransport);
    const client = new Client({ name: "test", version: "0.0.0" });
    await client.connect(clientTransport);
    const before = Date.now();
    await client.callTool({ name: "probe", arguments: {} });
    expect(deadline).toBeGreaterThan(before);

    // Used later, as a late click on a send confirmation would.
    await late!.get("/me");
    expect(seen[0]).toMatchObject({ signOutEpoch: 7 });
    expect(seen[0]).not.toHaveProperty("deadline");
    expect(seen[0]).not.toHaveProperty("signal");
  });
});
