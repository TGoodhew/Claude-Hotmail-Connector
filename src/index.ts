/**
 * Entry point — a local stdio MCP server (Mode A).
 *
 * Subcommands:
 *   (default)  start the MCP server on stdio (for Claude Desktop / VS Code)
 *   setup      auto-configure detected MCP hosts (Claude Desktop) to launch this server;
 *              `setup --client-id <id>` also saves the user's own Entra client id
 *   unsetup    remove this server from detected MCP hosts (used on uninstall)
 *   login      run the one-time interactive Microsoft sign-in and cache tokens
 *   logout     clear the cached tokens
 *   whoami     print the signed-in account (diagnostic)
 *
 * stdout is reserved exclusively for MCP JSON-RPC; all human/log output and the
 * sign-in UX go to stderr and the system browser.
 */

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createMicrosoftAuth } from "./auth/microsoft.js";
import { createFileTokenStore } from "./auth/tokens.js";
import { getConfig, MISSING_CLIENT_ID_MESSAGE, resolveCacheDir } from "./config.js";
import { createGraphClient } from "./graph/client.js";
import { whoami } from "./graph/user.js";
import { createMcpServer } from "./mcp.js";
import { isValidClientId, readSettings, writeSettings } from "./settings.js";
import { runSetup, runUnsetup } from "./setup/run.js";
import { ConfigError } from "./util/errors.js";
import { createLogger } from "./util/log.js";

/** Value of `--client-id <id>` / `--client-id=<id>`, if present. */
function clientIdFlag(args: string[]): string | undefined {
  for (let i = 0; i < args.length; i++) {
    const a = args[i]!;
    if (a === "--client-id") return args[i + 1] ?? "";
    if (a.startsWith("--client-id=")) return a.slice("--client-id=".length);
  }
  return undefined;
}

/**
 * `setup [--client-id <id>]`: save the client id (if given), then wire up the
 * host apps. Runs WITHOUT loading the full config, since setup is how the
 * client id gets configured in the first place.
 */
function setupCommand(args: string[]): void {
  const cacheDir = resolveCacheDir(process.env);
  const flag = clientIdFlag(args)?.trim();
  if (flag !== undefined) {
    if (!isValidClientId(flag)) {
      throw new ConfigError(
        `"${flag}" is not a valid Application (client) ID; expected a GUID like ` +
          "00000000-0000-0000-0000-000000000000.",
      );
    }
    const path = writeSettings(cacheDir, { clientId: flag });
    process.stderr.write(`Saved client id to ${path}\n`);
  }

  const clientId =
    flag || process.env.MICROSOFT_CLIENT_ID?.trim() || readSettings(cacheDir).clientId;
  if (!clientId) throw new ConfigError(MISSING_CLIENT_ID_MESSAGE);

  const outcome = runSetup({ clientId });
  for (const c of outcome.configured) {
    process.stderr.write(`  ${c.outcome}: ${c.path}\n`);
  }
  process.stderr.write(`\n${outcome.message}\n`);
}

async function run(): Promise<void> {
  const [command, ...args] = process.argv.slice(2);

  // Host wiring needs no Microsoft config, so it must not fail on it.
  if (command === "setup") {
    setupCommand(args);
    return;
  }

  if (command === "unsetup") {
    const outcome = runUnsetup();
    for (const c of outcome.configured) {
      process.stderr.write(`  removed: ${c.path}\n`);
    }
    process.stderr.write(`\n${outcome.message}\n`);
    return;
  }

  const log = createLogger({ name: "startup" });
  const config = getConfig();
  const store = createFileTokenStore(config.tokenCacheFile);
  const auth = createMicrosoftAuth({ config, store });
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

  // Default: start the MCP server. Do NOT block on interactive sign-in here —
  // that would stall the MCP handshake and pop a browser from a background
  // subprocess. Tools acquire tokens non-interactively; if not signed in they
  // return an actionable error telling the user to run `login`.
  const account = await auth.getAccount();
  if (!account) {
    log.warn("Not signed in yet. Run `node dist/index.js login` once to authenticate.");
  }

  const server = createMcpServer({ graph, config });
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
