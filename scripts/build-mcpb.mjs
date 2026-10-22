/**
 * Build the Claude Desktop extension (MCP Bundle, `.mcpb`): the connector's
 * distribution format for Claude Desktop (#63, spec v2 §11).
 *
 * An MCPB is a zip that Claude Desktop installs with a double-click. Claude
 * Desktop ships its own Node.js runtime, generates a settings UI from
 * `user_config`, and launches the server itself, so there's no installer, no
 * embedded runtime and no editing of claude_desktop_config.json.
 *
 * Steps:
 *   1. Bundle the app (incl. dependencies) into one CJS file with esbuild, so
 *      the bundle needs no node_modules.
 *   2. Write manifest.json (version from package.json).
 *   3. Validate and pack with the pinned official CLI (@anthropic-ai/mcpb).
 *
 * Run:  npm run build:mcpb   ->  build/hotmail-connector.mcpb
 */
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const root = process.cwd();
const pkg = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
const stageDir = join(root, "build", "mcpb");
const outFile = join(root, "build", "hotmail-connector.mcpb");

// One source of truth for the minimum Node: package.json engines (also needed by
// dependencies such as `open`). Claude Desktop checks the manifest against its own
// bundled Node, so the manifest must not claim older runtimes than we support.
const nodeRange = pkg.engines.node;
const nodeMajor = /([0-9]+)/.exec(nodeRange)?.[1];
if (!nodeMajor)
  throw new Error(`Cannot read a Node major version from engines.node "${nodeRange}"`);

function run(cmd, args) {
  console.log(`$ ${cmd} ${args.join(" ")}`);
  // A shell resolves `npx` (a .cmd on Windows); arguments here contain no spaces.
  execFileSync(cmd, args, { stdio: "inherit", shell: true });
}

rmSync(stageDir, { recursive: true, force: true });
mkdirSync(join(stageDir, "server"), { recursive: true });

// 1. Single-file server bundle.
run("npx", [
  "esbuild",
  "src/index.ts",
  "--bundle",
  "--platform=node",
  "--format=cjs",
  `--target=node${nodeMajor}`,
  "--outfile=build/mcpb/server/index.cjs",
]);
copyFileSync(join(root, "LICENSE"), join(stageDir, "LICENSE"));

// 2. Manifest (MCPB manifest spec 0.3).
const repo = "https://github.com/TGoodhew/Claude-Hotmail-Connector";
const manifest = {
  manifest_version: "0.3",
  name: "claude-hotmail-connector",
  display_name: "Hotmail Connector",
  version: pkg.version,
  description:
    "Read email and read/write calendar events in your personal Microsoft account " +
    "(Outlook.com / Hotmail / Live) via Microsoft Graph.",
  long_description:
    "A local MCP server for personal Microsoft accounts, which Claude's official Microsoft 365 " +
    "connector doesn't support. Mail is read-only; calendar writes are timezone-correct and " +
    "confirmation-gated, never add attendees, and deletion is a guarded soft delete. You need " +
    "your own Microsoft Entra app registration: see the README.",
  author: { name: "Tony Goodhew" },
  repository: { type: "git", url: `${repo}.git` },
  homepage: repo,
  documentation: `${repo}#readme`,
  support: `${repo}/blob/main/SUPPORT.md`,
  license: "MIT",
  keywords: ["outlook", "hotmail", "calendar", "email", "microsoft-graph"],
  server: {
    type: "node",
    entry_point: "server/index.cjs",
    mcp_config: {
      command: "node",
      args: ["${__dirname}/server/index.cjs"],
      env: { MICROSOFT_CLIENT_ID: "${user_config.client_id}" },
    },
  },
  user_config: {
    client_id: {
      type: "string",
      title: "Application (client) ID",
      description:
        "The Application (client) ID of your own Microsoft Entra app registration (README step 1, " +
        "or scripts/register-app.ps1). Not a secret.",
      required: true,
    },
  },
  tools: [
    { name: "whoami", description: "Show the connected personal Microsoft account." },
    { name: "list_messages", description: "Search or list email (read-only)." },
    { name: "get_message", description: "Read one email as text (read-only)." },
    { name: "get_attachment", description: "Read a text, HTML or .ics email attachment." },
    { name: "list_mail_folders", description: "List your mail folders (read-only)." },
    { name: "list_calendars", description: "List your calendars (own and shared)." },
    { name: "list_events", description: "List calendar events in a date range." },
    { name: "get_event", description: "Read one calendar event in full (notes, attendees)." },
    { name: "create_event", description: "Add a calendar event (timezone-correct, idempotent)." },
    { name: "update_event", description: "Edit a calendar event; reports before/after values." },
    { name: "cancel_event", description: "Guarded soft delete of a calendar event." },
    { name: "sign_out", description: "Remove the saved Microsoft sign-in from this computer." },
  ],
  compatibility: {
    platforms: ["win32", "darwin", "linux"],
    runtimes: { node: nodeRange },
  },
};
writeFileSync(join(stageDir, "manifest.json"), JSON.stringify(manifest, null, 2) + "\n");

// 3. Validate and pack.
run("npx", ["mcpb", "validate", "build/mcpb/manifest.json"]);
run("npx", ["mcpb", "pack", "build/mcpb", "build/hotmail-connector.mcpb"]);
console.log(`\n✓ Built ${outFile}`);
