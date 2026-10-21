# Claude Hotmail Connector

A **local** [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that lets Claude
**read email** and **read/write calendar events** in a **personal** Microsoft account
(`@hotmail.com` / `@outlook.com` / `@live.com`) via the Microsoft Graph API.

Claude's official Microsoft 365 connector only supports work/school (Microsoft Entra) accounts and
rejects personal Microsoft accounts. This project fills that gap as a **personal, single-user tool**
that runs entirely on your own machine.

> **Status:** v0.1.0 (Windows). The requirements are in [spec v2](docs/personal-outlook-mcp-connector-requirements.md),
> and the remaining work is tracked in the [Spec v2 epic (#66)](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/66).
> Where this README describes something that is about to change, it links the tracking issue.

## How it works — local stdio

The connector is a **local stdio MCP server** (spec v2 §4). Claude Desktop or Claude Code / VS Code
launches it as a subprocess and talks to it over stdin/stdout. There is **no public URL, no tunnel and
no hosting cost**, and no Claude ↔ server OAuth. The only sign-in is to **Microsoft**: a native
public-client OAuth 2.0 authorization-code + PKCE flow with a loopback redirect. No client secret is
involved.

```
Claude Desktop / Claude Code ──stdio──▶ this server ──HTTPS──▶ Microsoft Graph (personal account)
        (MCP client)                    (local subprocess)        graph.microsoft.com/v1.0
```

Because it runs locally, it works in **Claude Desktop** and **Claude Code / VS Code** on the same PC. It
does **not** work in claude.ai on the web or mobile; that would need a remote mode, which is a future
phase ([#58](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/58)).

## Tools

| Tool            | Purpose                                             | Graph permission      |
| --------------- | --------------------------------------------------- | --------------------- |
| `whoami`        | Confirm the connected account                       | `User.Read`           |
| `list_messages` | Search/list mail (find booking confirmations)       | `Mail.Read`           |
| `get_message`   | Read one email as text                              | `Mail.Read`           |
| `list_events`   | Read the calendar over a date range                 | `Calendars.Read`      |
| `create_event`  | Add a calendar event (timezone-correct, idempotent) | `Calendars.ReadWrite` |
| `update_event`  | Edit a calendar event (destructive: overwrites)     | `Calendars.ReadWrite` |

Mail is **read-only** (no `Mail.Send`). There is **no hard-delete** tool, and calendar writes never add
attendees or send invitations. `cancel_event` (event deletion) is undecided and gated
([#55](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/55)). Full tool contracts are in spec v2 §7.

## Install (Windows)

Download the latest `HotmailConnectorSetup.exe` (or the portable `hotmail-connector.exe`) from
[Releases](https://github.com/TGoodhew/Claude-Hotmail-Connector/releases) and run it. The installer
configures Claude Desktop and offers to run the one-time Microsoft sign-in. Verify downloads against
`SHA256SUMS.txt`.

- Windows shows an **"unknown publisher"** warning (the build is unsigned, by design). Click
  **More info → Run anyway**.
- After it finishes, **fully quit Claude Desktop from the system tray** (closing the window only
  minimises it) and reopen it.

> **Changing soon:** v0.1.0 embeds a shared public-client id, so no Entra registration is needed. Per
> spec v2 this is a single-user tool, so that shared id is being removed: you'll create **your own**
> Entra app registration (step 1 below) and the installer will ask for its client id
> ([#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68)). Packaging as a
> one-click Claude Desktop extension (`.mcpb`) is being evaluated
> ([#63](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/63)).

See [`docs/installer.md`](docs/installer.md) for installer details and how the packages are built.

## Manual / developer setup

### Prerequisites

- **Node.js 20+**. **Node 24 LTS is recommended**: Node 20 reached end-of-life in April 2026, and
  the project is moving to 24 ([#45](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/45)).
- A **personal Microsoft account** (`@hotmail.com` / `@outlook.com` / `@live.com`).
- **Claude Desktop** and/or **Claude Code / VS Code** (with MCP support) on the same machine.
- A **Microsoft Entra app registration**, which is free (step 1). Registering an app needs a
  **Microsoft Entra tenant**, and a personal account doesn't come with one: if you don't have a
  tenant, [create a free one](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-create-new-tenant)
  first (typically by signing up for a free Azure account).

### 1. Register a Microsoft Entra application

> Optional in v0.1.0 (the bundled shared client id is used if you skip it), and **required** once
> [#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68) lands. If you have the Azure
> CLI signed in to your tenant (`az login`), `scripts/register-app.ps1` automates these steps.

The connector signs in as a **public/native client**. No client secret is ever used or stored.

1. Go to the [Entra admin center → App registrations](https://entra.microsoft.com/) →
   **New registration**.
2. **Name:** e.g. `Claude Hotmail Connector`.
3. **Supported account types:** choose **"Accounts in any organizational directory (Any Microsoft
   Entra ID tenant – Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)."** This is what
   allows `hotmail.com` / `outlook.com` sign-in with the default `common` tenant. (Restricting this to
   personal accounts only is under review in
   [#62](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/62).)
4. **Redirect URI:** select platform **"Public client/native (mobile & desktop)"** and enter
   **`http://localhost`**. Microsoft allows any loopback port for native clients, so you don't need a
   specific port.
5. Click **Register**, then copy the **Application (client) ID**.
6. Under **Authentication**, confirm the platform is **Mobile and desktop applications** with
   `http://localhost`, and that **"Allow public client flows"** is **Yes**. Do **not** create a
   client secret.
7. Under **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add:
   `openid`, `profile`, `email`, `offline_access`, `User.Read`, `Mail.Read`, `Calendars.ReadWrite`.
   A personal account doesn't need admin consent; you consent at first sign-in.

### 2. Install & build

```bash
git clone https://github.com/TGoodhew/Claude-Hotmail-Connector.git
cd Claude-Hotmail-Connector
npm install
npm run build          # produces dist/index.js
```

### 3. Configure

Set your client id (step 1) in a local `.env` for development, or in the client's `env` block when
wiring it up (step 5):

```bash
cp .env.example .env
# edit .env and set MICROSOFT_CLIENT_ID=<the Application (client) ID from step 1>
```

Environment variables (see [`.env.example`](.env.example)):

| Variable                   | Default                       | Notes                                                                                                                           |
| -------------------------- | ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `MICROSOFT_CLIENT_ID`      | bundled shared id (v0.1.0)    | Your Entra app's client id. Becomes required ([#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68)).           |
| `MICROSOFT_TENANT`         | `common`                      | `common` allows personal + work/school accounts.                                                                                |
| `DEFAULT_TIMEZONE`         | `Australia/Brisbane`          | IANA zone used when a call doesn't specify one ([#34](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/34)).         |
| `LOG_LEVEL`                | `info`                        | `debug` / `info` / `warn` / `error` (logs go to stderr).                                                                        |
| `MICROSOFT_SCOPES`         | least-privilege set (spec §6) | Space/comma-separated override; becomes narrowing-only ([#70](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/70)). |
| `CLAUDE_HOTMAIL_CACHE_DIR` | per-OS profile dir            | Where the encrypted token cache lives.                                                                                          |

### 4. Sign in once

```bash
node dist/index.js login
```

This opens your system browser to Microsoft. Sign in with your **personal** account and consent to
the requested permissions. The refresh token is stored **encrypted** in your user profile (see
[Security](#security)), and later runs refresh it silently. `node dist/index.js logout` clears it.

Verify:

```bash
node dist/index.js whoami       # prints your name and account, e.g. Jane Doe <jane@outlook.com>
```

### 5. Wire it into a client

The quickest way is to let the connector configure Claude Desktop for you (the same engine the
installer uses):

```bash
node dist/index.js setup     # detects Claude Desktop (incl. the Microsoft Store build) and adds the server
node dist/index.js unsetup   # removes it again
```

Or configure a client manually. The server is launched as a stdio subprocess; pass your
`MICROSOFT_CLIENT_ID` in its `env`.

#### Claude Code / VS Code

Use **Command Palette → "MCP: Add Server…" → Command (stdio)**, pointing `node` at the built
`dist/index.js`. Or copy [`.vscode/mcp.json.example`](.vscode/mcp.json.example) to `.vscode/mcp.json`
and edit the path and client id:

```jsonc
{
  "servers": {
    "hotmail-connector": {
      "type": "stdio",
      "command": "node",
      "args": ["C:\\path\\to\\Claude-Hotmail-Connector\\dist\\index.js"],
      "env": { "MICROSOFT_CLIENT_ID": "<your-entra-app-client-id>" },
    },
  },
}
```

#### Claude Desktop

Edit `claude_desktop_config.json` (see
[`examples/claude_desktop_config.example.json`](examples/claude_desktop_config.example.json)):

```jsonc
{
  "mcpServers": {
    "hotmail-connector": {
      "command": "node",
      "args": ["C:\\path\\to\\Claude-Hotmail-Connector\\dist\\index.js"],
      "env": { "MICROSOFT_CLIENT_ID": "<your-entra-app-client-id>" },
    },
  },
}
```

Restart the client and the connector's tools should appear. If you haven't run `login` yet, run it
once in a terminal first: the server doesn't open a browser on its own today. On-demand sign-in is
planned in [#69](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/69).

## Usage example

> "Find my Monsoon Aquatics and Macadamias Australia booking confirmation emails, then add each to my
> calendar with 15-minute travel-time blocks before and after."

Claude uses `list_messages` / `get_message` to read the confirmations and confirms the details with
you. It then calls `create_event` for each booking (and the travel blocks) at the correct local times,
with no duplicates if it retries.

## Security

- **Least privilege:** `Mail.Read` (read-only) + `Calendars.ReadWrite`. No `Mail.Send`, no hard
  delete, no attendees/invitations.
- **No client secret:** a public/native client using PKCE.
- **Token storage:** the refresh token is encrypted with AES-256-GCM in your user profile
  (`%LOCALAPPDATA%\claude-hotmail-connector` on Windows, `~/.claude-hotmail-connector` otherwise). The
  key is currently in a sibling file in the same folder, which isn't strong protection on its own. Spec
  v2 requires OS-protected storage (DPAPI / Keychain / libsecret), tracked in
  [#42](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/42). Tokens are never committed,
  logged or returned to the client.
- **Logging:** structured logs go to **stderr only** (stdout is reserved for MCP JSON-RPC) and are
  scrubbed of tokens, `Authorization` headers, OAuth codes and secrets.
- **Email is untrusted input:** messages are written by strangers and may try to steer the model. The
  threat model and its mitigations are in spec v2 §9
  ([#60](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/60)). Review what Claude proposes
  before approving calendar writes.
- **Timezone-correct writes:** events are sent as local wall-clock time + an IANA time zone (never a
  UTC `Z` with a named zone), so they land at the intended local time.

Known security issues are tracked under the `area:security` label.

## Troubleshooting

- **`AADSTS50020` / "account does not exist in tenant" / personal account rejected:** the app
  registration isn't set to **multitenant + personal accounts**, or you used a tenant other than
  `common`. Recreate or adjust it per step 1.3 and keep `MICROSOFT_TENANT=common`.
- **`AADSTS7000218` / "client_assertion or client_secret required":** the app is registered as a
  confidential/web client. Register it as a **public client** with **"Allow public client flows" =
  Yes** (step 1.6).
- **`redirect_uri` mismatch:** add **`http://localhost`** under the Mobile & desktop platform.
- **Event lands an hour off:** always pass a local wall-clock `dateTime` (e.g. `2026-07-20T10:00:00`)
  with an IANA `timeZone`, never a `Z`-suffixed value. The connector rejects `Z`/offset values for
  event times.
- **"Not signed in" from a tool:** run `node dist/index.js login` (or the **Sign in to Hotmail
  Connector** Start-menu shortcut) once. If it persists, run `logout` then `login` to reset the token
  cache.
- **`npm install` times out or hangs (dead IPv6 route):** force IPv4:
  `NODE_OPTIONS="--dns-result-order=ipv4first --no-network-family-autoselection" npm install`.

## Development

```bash
npm run dev             # rebuild on change (tsup --watch)
npm run typecheck       # tsc --noEmit
npm run lint            # eslint
npm run format          # prettier --write   (format:check in CI)
npm run check:text      # reject raw control characters in tracked files (CI)
npm test                # vitest (unit + in-memory + built-subprocess smoke)
npm run build:exe       # Windows: Node SEA single executable (postject via npx)
npm run build:installer # Windows: Inno Setup installer (needs iscc on PATH)
```

Layout (tests are colocated as `src/**/*.test.ts`):

- `src/index.ts`: CLI and stdio entry (`setup`, `unsetup`, `login`, `logout`, `whoami`; default =
  server). `src/mcp.ts`: server assembly.
- `src/auth`: Microsoft OAuth (PKCE + loopback) and the encrypted token cache.
- `src/graph`: Graph client (auth, retries, paging) plus the mail / calendar / user modules.
- `src/tools`: zod schemas and thin tool handlers.
- `src/setup`: MCP host auto-configuration behind `setup` / `unsetup`.
- `src/util`: time zones, HTML-to-text, logging, errors.

## Requirements, design & roadmap

- [`docs/personal-outlook-mcp-connector-requirements.md`](docs/personal-outlook-mcp-connector-requirements.md):
  the build spec, **v2** (tools, Graph, auth, security, timezone, packaging, decision log).
  Requirements not yet implemented link to their tracking issues.
- [`docs/local-hosting-investigation-and-requirements.md`](docs/local-hosting-investigation-and-requirements.md):
  the original local-hosting investigation (superseded by spec v2; kept for the record).
- [`docs/installer.md`](docs/installer.md): the packaged Windows installer, the auto-config engine,
  `register-app`, and the SEA/Inno build.
- **Roadmap:** the [Spec v2 epic (#66)](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/66)
  orders all open work: spec decisions → security fixes → correctness → CI/tests → features.

## License

[MIT](LICENSE) © 2026 Tony Goodhew
