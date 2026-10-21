/**
 * Command-line dispatch for the entry point (#39).
 *
 * With no arguments the process is an MCP stdio server, as clients launch it.
 * Anything else must be a known command: a typo or `--help` used to start the
 * server silently, which looks like a hang at a terminal.
 */

import { SERVER_NAME, SERVER_VERSION } from "./mcp.js";

export type Command =
  | { kind: "server" }
  | { kind: "login" | "logout" | "whoami" }
  | { kind: "help" | "version" }
  | { kind: "invalid"; message: string };

export const USAGE = `${SERVER_NAME} ${SERVER_VERSION}

Usage: node dist/index.js [command]

With no command, runs the MCP server on stdio. MCP clients (Claude Desktop,
Claude Code, VS Code) start it this way; there's no need to run it yourself.

Commands:
  login      Sign in to your personal Microsoft account and save the sign-in
  logout     Delete the saved sign-in (token cache and its key)
  whoami     Show the signed-in account
  help       Show this help (also -h, --help)
  version    Show the version (also -v, --version)
`;

const ACCOUNT_COMMANDS = new Set(["login", "logout", "whoami"]);

/** Interpret the arguments after the script path (`process.argv.slice(2)`). */
export function parseCommand(args: readonly string[]): Command {
  if (args.length === 0) return { kind: "server" };
  const [first, ...rest] = args;
  if (rest.length > 0) {
    return { kind: "invalid", message: `Unexpected argument: ${rest[0]}` };
  }
  if (first === "help" || first === "-h" || first === "--help") return { kind: "help" };
  if (first === "version" || first === "-v" || first === "--version") return { kind: "version" };
  if (ACCOUNT_COMMANDS.has(first!)) return { kind: first as "login" | "logout" | "whoami" };
  return { kind: "invalid", message: `Unknown command: ${first}` };
}
