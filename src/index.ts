/**
 * Entry point — a local stdio MCP server (Mode A).
 *
 * Normally launched by an MCP client: by Claude Desktop as an extension
 * (`.mcpb`, see scripts/build-mcpb.mjs), or by Claude Code / VS Code from a
 * source checkout. The client id comes from `MICROSOFT_CLIENT_ID` (the
 * extension's "Application (client) ID" setting, or the client's `env`).
 *
 * Subcommands (for source checkouts):
 *   (default)  start the MCP server on stdio
 *   login      run the interactive Microsoft sign-in and cache tokens
 *   logout     clear the cached tokens
 *   whoami     print the signed-in account (diagnostic)
 *
 * stdout is reserved exclusively for MCP JSON-RPC; all human/log output and the
 * sign-in UX go to stderr and the system browser.
 */

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createMicrosoftAuth } from "./auth/microsoft.js";
import { defaultKeyProtector } from "./auth/protect.js";
import { createFileTokenStore } from "./auth/tokens.js";
import { getConfig } from "./config.js";
import { createGraphClient } from "./graph/client.js";
import { whoami } from "./graph/user.js";
import { createMcpServer } from "./mcp.js";
import { createLogger } from "./util/log.js";

/** How long a tool call waits for an on-demand browser sign-in before asking the user to retry. */
const ON_DEMAND_SIGN_IN_WAIT_MS = 45_000;

async function run(): Promise<void> {
  const command = process.argv[2];
  const log = createLogger({ name: "startup" });
  const config = getConfig();
  const store = createFileTokenStore(config.tokenCacheFile, { protector: defaultKeyProtector() });
  // At a terminal (`login`), a sign-in URL the user can copy beats failing.
  const auth = createMicrosoftAuth({ config, store, printUrlIfBrowserFails: command === "login" });
  const graph = createGraphClient({ auth, baseUrl: config.graphBaseUrl });

  if (command === "login") {
    const tokens = await auth.signIn();
    log.info("Signed in.", { account: tokens.account });
    process.stderr.write(`Signed in as ${tokens.account ?? "your account"}.\n`);
    return;
  }

  if (command === "logout") {
    await auth.signOut();
    process.stderr.write("Signed out. Cached tokens cleared.\n");
    return;
  }

  if (command === "whoami") {
    const me = await whoami(graph);
    process.stderr.write(`${me.displayName ?? ""} <${me.mail ?? me.userPrincipalName ?? "?"}>\n`);
    return;
  }

  // Default: start the MCP server. Never sign in at startup: that would stall
  // the MCP handshake and pop a browser before the user has asked for
  // anything. Instead, the first tool call that needs a token starts the
  // browser sign-in on demand and waits up to ON_DEMAND_SIGN_IN_WAIT_MS (inside
  // typical 60s MCP request timeouts). Past that, the call returns "finish
  // signing in, then retry" while the sign-in carries on in the background.
  const account = await auth.getAccount();
  if (!account) {
    log.warn("Not signed in yet. The first tool call will open a browser sign-in.");
  }

  const toolGraph = createGraphClient({
    auth,
    baseUrl: config.graphBaseUrl,
    interactiveSignIn: { waitMs: ON_DEMAND_SIGN_IN_WAIT_MS },
  });
  const server = createMcpServer({ graph: toolGraph, config, auth });
  const transport = new StdioServerTransport();
  await server.connect(transport);
  log.info("MCP server ready on stdio.", { account });

  const shutdown = (): void => {
    void server.close().finally(() => process.exit(0));
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}

run().catch((err: unknown) => {
  // Startup failures go to stderr; stdout stays clean.
  process.stderr.write(`Fatal: ${err instanceof Error ? err.message : String(err)}\n`);
  process.exit(1);
});
