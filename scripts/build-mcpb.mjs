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
 *      the bundle needs no node_modules; likewise the PDF worker (#156). List
 *      every bundled package, with its license, in THIRD_PARTY_NOTICES.txt.
 *   2. Write manifest.json (version from package.json).
 *   3. Validate and pack with the pinned official CLI (@anthropic-ai/mcpb).
 *
 * Run:  npm run build:mcpb   ->  build/hotmail-connector.mcpb
 */
import { execFileSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
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

// 1. Single-file server bundle, and the PDF worker it starts beside it (#156).
// The server picks the worker's path by format: as CommonJS it has __dirname, so
// the ESM-only import.meta branch is never reached here, hence the silenced warning.
const bundled = new Set();
for (const [entry, out] of [
  ["src/index.ts", "index.cjs"],
  ["src/pdf-worker.ts", "pdf-worker.cjs"],
]) {
  const metafile = join(root, "build", `${out}.meta.json`);
  run("npx", [
    "esbuild",
    entry,
    "--bundle",
    "--platform=node",
    "--format=cjs",
    `--target=node${nodeMajor}`,
    "--log-override:empty-import-meta=silent",
    `--metafile=${metafile}`,
    `--outfile=build/mcpb/server/${out}`,
  ]);
  // The package directory of every bundled input under node_modules.
  for (const input of Object.keys(JSON.parse(readFileSync(metafile, "utf8")).inputs)) {
    const at = input.lastIndexOf("node_modules/");
    if (at < 0) continue;
    const parts = input.slice(at + "node_modules/".length).split("/");
    const name = parts[0].startsWith("@") ? `${parts[0]}/${parts[1]}` : parts[0];
    bundled.add(input.slice(0, at) + "node_modules/" + name);
  }
  rmSync(metafile);
}

// Licenses of what's bundled: the MIT and Apache licenses ask for their notices
// to travel with the code. pdf.js comes inside unpdf, without its own license
// file, so that one is kept in licenses/.
const notices = [
  "Third-party software bundled in the Hotmail Connector extension.",
  "The connector itself is MIT-licensed: see LICENSE.",
  "",
];
for (const dir of [...bundled].sort()) {
  const pkgJson = JSON.parse(readFileSync(join(root, dir, "package.json"), "utf8"));
  const licenseFile = readdirSync(join(root, dir)).find((f) => /^(licen[cs]e|copying)/i.test(f));
  notices.push(
    "=".repeat(78),
    `${pkgJson.name} ${pkgJson.version} (${pkgJson.license ?? "see below"})`,
    "=".repeat(78),
    licenseFile
      ? readFileSync(join(root, dir, licenseFile), "utf8").trim()
      : "(no license file in the package)",
    "",
  );
  if (pkgJson.name === "unpdf") {
    const pdfjs = join(root, "licenses", "pdf.js-LICENSE.txt");
    if (!existsSync(pdfjs)) throw new Error("licenses/pdf.js-LICENSE.txt is missing");
    notices.push(
      "=".repeat(78),
      "pdf.js (bundled inside unpdf), Copyright Mozilla Foundation (Apache-2.0)",
      "=".repeat(78),
      readFileSync(pdfjs, "utf8").trim(),
      "",
    );
  }
}
writeFileSync(join(stageDir, "THIRD_PARTY_NOTICES.txt"), notices.join("\n"));
console.log(`Listed ${bundled.size} bundled packages in THIRD_PARTY_NOTICES.txt`);
copyFileSync(join(root, "LICENSE"), join(stageDir, "LICENSE"));
// The icon Claude Desktop shows for the extension (#179): a 512×512 PNG rendered
// from assets/icon.svg (see docs/packaging.md).
copyFileSync(join(root, "assets", "icon.png"), join(stageDir, "icon.png"));

// 2. Manifest (MCPB manifest spec 0.3).
const repo = "https://github.com/TGoodhew/Claude-Hotmail-Connector";
const manifest = {
  manifest_version: "0.3",
  name: "claude-hotmail-connector",
  display_name: "Hotmail Connector",
  version: pkg.version,
  description:
    "Read email and read/write calendar events in your personal Microsoft account " +
    "(Outlook.com / Hotmail / Live) via Microsoft Graph, and optionally send email you confirm.",
  long_description:
    "A local MCP server for personal Microsoft accounts, which Claude's official Microsoft 365 " +
    "connector doesn't support. Mail is read-only unless you turn on sending, and then every " +
    "email is shown to you in your browser and sent only when you click Send. Calendar writes " +
    "are timezone-correct and confirmation-gated, never add attendees, and deletion is a guarded " +
    "soft delete. You need your own Microsoft Entra app registration: see the README.",
  author: { name: "Tony Goodhew" },
  repository: { type: "git", url: `${repo}.git` },
  homepage: repo,
  documentation: `${repo}#readme`,
  support: `${repo}/blob/main/SUPPORT.md`,
  license: "MIT",
  icon: "icon.png",
  keywords: ["outlook", "hotmail", "calendar", "email", "microsoft-graph"],
  server: {
    type: "node",
    entry_point: "server/index.cjs",
    mcp_config: {
      command: "node",
      args: ["${__dirname}/server/index.cjs"],
      env: {
        MICROSOFT_CLIENT_ID: "${user_config.client_id}",
        ALLOW_SEND_MAIL: "${user_config.allow_send}",
        ALLOW_MAIL_DRAFTS: "${user_config.allow_drafts}",
        LOG_FILE: "${user_config.log_file}",
      },
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
    // Sending mail (#57): off unless the user turns it on.
    allow_send: {
      type: "boolean",
      title: "Allow sending email",
      description:
        "Let Claude send emails and replies from your account. Each one is shown to you in your " +
        "browser first, and sent only if you click Send there. Turning this on asks you to sign " +
        "in to Microsoft once more, to allow sending.",
      default: false,
      required: false,
    },
    // Saving drafts (#184): off unless the user turns it on.
    allow_drafts: {
      type: "boolean",
      title: "Allow saving email drafts",
      description:
        "Let Claude save emails and replies as drafts in your Drafts folder and open them in " +
        "Outlook, where you review, edit and send them yourself. Claude can't send them. This " +
        "gives the connector permission to change and delete mail, which it uses only to create " +
        "drafts. Turning this on asks you to sign in to Microsoft once more.",
      default: false,
      required: false,
    },
    // Claude Desktop doesn't keep the extension's stderr, so logs can go to a file (#178).
    log_file: {
      type: "boolean",
      title: "Write a log file",
      description:
        "Save the connector's logs to connector.log, to help diagnose problems. It never holds " +
        "tokens, your email address, or what's in your emails and events, and never grows past " +
        "2 MB. It's in " +
        "%LOCALAPPDATA%\\Packages\\Claude_pzs8sxrjxfjjc\\LocalCache\\Local\\claude-hotmail-connector " +
        "(see the README).",
      default: false,
      required: false,
    },
  },
  tools: [
    { name: "whoami", description: "Show the connected personal Microsoft account." },
    { name: "list_messages", description: "Search or list email (read-only)." },
    { name: "get_message", description: "Read one email as text (read-only)." },
    {
      name: "send_email",
      description: "Send an email, after you confirm it in your browser (if sending is on).",
    },
    {
      name: "reply_to_email",
      description: "Reply to an email's sender, after you confirm it (if sending is on).",
    },
    {
      name: "create_draft",
      description: "Save an email or reply as a draft and open it in Outlook (if drafts are on).",
    },
    { name: "get_attachment", description: "Read a text, HTML, PDF or .ics email attachment." },
    { name: "list_mail_folders", description: "List your mail folders (read-only)." },
    { name: "list_calendars", description: "List your calendars (own and shared)." },
    { name: "list_events", description: "List calendar events in a date range." },
    { name: "get_event", description: "Read one calendar event in full (notes, attendees)." },
    { name: "find_free_time", description: "Find free time in a date range." },
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
