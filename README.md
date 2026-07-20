# Claude Hotmail Connector

A **locally hosted** [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that lets
Claude read email and read/write calendar events in a **personal** Microsoft account
(`@hotmail.com` / `@outlook.com` / `@live.com`) via the Microsoft Graph API.

Claude's official Microsoft 365 connector only supports work/school (Microsoft Entra) accounts and
rejects personal Microsoft accounts. This project fills that gap for a single user, running entirely
on your own machine.

## Deployment model — Mode A: local stdio

Per the [local-hosting investigation](docs/local-hosting-investigation-and-requirements.md), this
connector is built as a **local stdio MCP server**: Claude Desktop or VS Code / Claude Code launches
it as a subprocess and talks to it over stdin/stdout. There is **no public URL, no tunnel, and no
hosting cost** — and the entire downstream OAuth machinery (PRM/DCR/PKCE for the Claude ↔ server hop)
drops away. Only the **upstream** Microsoft sign-in is needed, done via a native public-client
OAuth 2.0 authorization-code + PKCE flow with a loopback redirect.

```
Claude Desktop / VS Code ──stdio──▶ this server ──HTTPS──▶ Microsoft Graph (personal account)
        (MCP client)                (local subprocess)        graph.microsoft.com/v1.0
```

## Tools

| Tool            | Purpose                                             | Graph permission      |
| --------------- | --------------------------------------------------- | --------------------- |
| `whoami`        | Confirm the connected account                       | `User.Read`           |
| `list_messages` | Search/list mail (find booking confirmations)       | `Mail.Read`           |
| `get_message`   | Read one email as text                              | `Mail.Read`           |
| `list_events`   | Read the calendar over a date range                 | `Calendars.Read`      |
| `create_event`  | Add a calendar event (timezone-correct, idempotent) | `Calendars.ReadWrite` |
| `update_event`  | Edit a calendar event                               | `Calendars.ReadWrite` |

Mail is **read-only** (no `Mail.Send`). There is **no hard-delete** tool. `cancel_event` (event
deletion) is deliberately **deferred** as a gated/destructive action pending explicit approval.

## Install (recommended)

The easy path — **no Node, no Azure, no config files**. On **Windows**, download the latest
`HotmailConnectorSetup.exe` (or the portable `hotmail-connector.exe`) from
[Releases](https://github.com/TGoodhew/Claude-Hotmail-Connector/releases) and run it. The installer
auto-configures Claude Desktop and runs the one-time Microsoft sign-in for you — the build embeds a
shared public-client id, so there's nothing to register or configure.

- Windows shows an **"unknown publisher"** warning (the build is unsigned, by design) — click
  **More info → Run anyway**.
- After it finishes, **fully quit Claude Desktop from the system tray** (closing the window only
  minimises it) and reopen it.

See [`docs/installer.md`](docs/installer.md) for the details, the SmartScreen note, and how the
packages are built. Want to build or run it yourself instead? Use the developer setup below.

## Manual / developer setup

### Prerequisites

- **Node.js 20+** (`node --version`).
- A **personal Microsoft account** (`@hotmail.com` / `@outlook.com` / `@live.com`).
- **Claude Desktop** and/or **VS Code** (with MCP support) on the same machine.
- A one-time **Microsoft Entra app registration** (free — see below).

### 1. Register a Microsoft Entra application (one-time, optional)

> You can **skip this** — the build ships with a bundled shared client id. Only do it to bring your
> own Entra registration. (Maintainers: `scripts/register-app.ps1` automates it; see
> [`docs/installer.md`](docs/installer.md).)

The connector signs in as a **public/native client** — no client secret is ever used or stored.

1. Go to the [Entra admin center → App registrations](https://entra.microsoft.com/) →
   **New registration**.
2. **Name:** e.g. `Claude Hotmail Connector`.
3. **Supported account types:** choose **“Accounts in any organizational directory (Any Microsoft
   Entra ID tenant – Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox).”**
   This is what allows `hotmail.com` / `outlook.com` sign-in.
4. **Redirect URI:** select platform **“Public client/native (mobile & desktop)”** and enter
   **`http://localhost`**. (Microsoft allows any loopback port for native clients, so a specific port
   is not required.)
5. Click **Register**, then copy the **Application (client) ID**.
6. Under **Authentication**, confirm the platform is **Mobile and desktop applications** with
   `http://localhost`, and that **“Allow public client flows”** is **Yes**. Do **not** create a
   client secret.
7. Under **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add:
   `openid`, `profile`, `email`, `offline_access`, `User.Read`, `Mail.Read`, `Calendars.ReadWrite`.
   (Admin consent is not required for a personal account — you consent at first sign-in.)

### 2. Install & build

```bash
git clone https://github.com/TGoodhew/Claude-Hotmail-Connector.git
cd Claude-Hotmail-Connector
npm install
npm run build          # produces dist/index.js
```

### 3. Configure (optional)

The build ships with a bundled shared client id, so this is **optional**. To use your own Entra
registration (step 1), set it via a local `.env`:

```bash
cp .env.example .env
# edit .env and set MICROSOFT_CLIENT_ID=<the Application (client) ID from step 1>
```

Configurable environment variables (see [`.env.example`](.env.example)):

| Variable                   | Default              | Notes                                           |
| -------------------------- | -------------------- | ----------------------------------------------- |
| `MICROSOFT_CLIENT_ID`      | _(required)_         | Entra app registration client id.               |
| `MICROSOFT_TENANT`         | `common`             | `common` enables personal + work accounts.      |
| `DEFAULT_TIMEZONE`         | `Australia/Brisbane` | IANA zone for calendar reads/writes.            |
| `LOG_LEVEL`                | `info`               | `debug`/`info`/`warn`/`error` (logs to stderr). |
| `MICROSOFT_SCOPES`         | least-privilege set  | Space/comma-separated override.                 |
| `CLAUDE_HOTMAIL_CACHE_DIR` | per-OS profile dir   | Where the encrypted token cache lives.          |

### 4. Sign in once

```bash
node dist/index.js login
```

This opens your system browser to Microsoft. Sign in with your **personal** account and consent to
the requested permissions. The refresh token is stored **encrypted** in your user profile (see
[Security](#security)); subsequent runs refresh silently. `node dist/index.js logout` clears it.

Verify:

```bash
node dist/index.js whoami       # should print your name <tony_goodhew@hotmail.com>
```

### 5. Wire it into a client

The quickest way is to let the connector configure Claude Desktop for you — the same engine the
installer uses:

```bash
node dist/index.js setup   # detects Claude Desktop (incl. the Microsoft Store build) and wires it up
```

Or configure a client manually — the server is launched as a stdio subprocess. Pass
`MICROSOFT_CLIENT_ID` via the client's `env` only if you're bringing your own registration (otherwise
the bundled default is used).

### VS Code / Claude Code

Use **Command Palette → “MCP: Add Server…” → Command (stdio)**, pointing `node` at the built
`dist/index.js`. Or commit-adjacent, copy [`.vscode/mcp.json.example`](.vscode/mcp.json.example) to
`.vscode/mcp.json` and edit the path/client id:

```jsonc
{
  "servers": {
    "hotmail-connector": {
      "type": "stdio",
      "command": "node",
      "args": ["C:\\Users\\Tony\\Source\\Repos\\Claude-Hotmail-Connector\\dist\\index.js"],
      "env": { "MICROSOFT_CLIENT_ID": "<your-entra-app-client-id>" },
    },
  },
}
```

### Claude Desktop

Edit `claude_desktop_config.json` (see
[`examples/claude_desktop_config.example.json`](examples/claude_desktop_config.example.json)):

```jsonc
{
  "mcpServers": {
    "hotmail-connector": {
      "command": "node",
      "args": ["C:\\Users\\Tony\\Source\\Repos\\Claude-Hotmail-Connector\\dist\\index.js"],
      "env": { "MICROSOFT_CLIENT_ID": "<your-entra-app-client-id>" },
    },
  },
}
```

Restart the client. The connector's tools should appear. If you have not run `login` yet, do so once
in a terminal first (the server does not pop a browser on its own).

## Usage example

> “Find my Monsoon Aquatics and Macadamias Australia booking confirmation emails, then add each to my
> calendar with 15-minute travel-time blocks before and after.”

Claude will use `list_messages` / `get_message` to read the confirmations, confirm the details with
you, then call `create_event` for each booking (and the travel blocks) at the correct
`Australia/Brisbane` local times — with no duplicates if it retries.

## Security

- **Least privilege:** `Mail.Read` (read-only) + `Calendars.ReadWrite`. No `Mail.Send`, no
  hard-delete, `cancel_event` deferred.
- **No client secret:** public/native client using PKCE.
- **Token storage:** the refresh token is encrypted with AES-256-GCM in your user profile
  (`%LOCALAPPDATA%\claude-hotmail-connector` on Windows, `~/.claude-hotmail-connector` otherwise),
  with the key in a sibling owner-only file. Never committed, never logged, never returned to the
  client.
- **Logging:** structured logs go to **stderr only** (stdout is reserved for MCP JSON-RPC) and are
  scrubbed of tokens, `Authorization` headers, OAuth codes, and secrets.
- **Timezone-correct writes:** events are sent as local wall-clock + IANA time zone (never a UTC `Z`
  with a named zone), so they land at the intended local time.

## Troubleshooting

- **`AADSTS50020` / “account does not exist in tenant” / personal-account rejected:** the app
  registration is not set to **multitenant + personal accounts**, or you used a tenant other than
  `common`. Recreate/adjust per step 1.3 and keep `MICROSOFT_TENANT=common`.
- **`AADSTS7000218` / “client_assertion or client_secret required”:** the app is registered as a
  confidential/web client. Register it as a **public client** with **“Allow public client flows” =
  Yes** (step 1.6).
- **`redirect_uri` mismatch:** add **`http://localhost`** under the Mobile & desktop platform.
- **Event lands an hour off:** always pass local wall-clock `dateTime` (e.g. `2026-07-20T10:00:00`)
  with an IANA `timeZone`; never a `Z`-suffixed value. The connector rejects `Z`/offset values for
  event times.
- **“Not signed in” from a tool:** run `node dist/index.js login` once; if it persists, `logout`
  then `login` again to reset the token cache.
- **`npm install` times out / hangs (dead IPv6 route):** force IPv4 —
  `NODE_OPTIONS="--dns-result-order=ipv4first --no-network-family-autoselection" npm install`.

## Development

```bash
npm run dev             # rebuild on change (tsup --watch)
npm run typecheck       # tsc --noEmit
npm run lint            # eslint
npm run format          # prettier --write
npm test                # vitest (unit + in-memory + built-subprocess smoke)
npm run build:exe       # Windows: Node SEA single executable (postject via npx)
npm run build:installer # Windows: Inno Setup installer (needs iscc on PATH)
```

Layout: `src/auth` (Microsoft OAuth + encrypted token cache), `src/graph` (Graph client + mail /
calendar / user modules), `src/tools` (zod schemas + thin handlers), `src/setup` (client auto-config

- the `setup`/`unsetup` commands), `src/util` (time, logging, errors, html), `src/mcp.ts` +
  `src/index.ts` (server assembly + stdio entry).

## Requirements & design

The authoritative specifications live in [`docs/`](docs/):

- [`docs/personal-outlook-mcp-connector-requirements.md`](docs/personal-outlook-mcp-connector-requirements.md) — the main build spec (tools, Graph, security, timezone).
- [`docs/local-hosting-investigation-and-requirements.md`](docs/local-hosting-investigation-and-requirements.md) — why/how to host locally (Mode A stdio, chosen here).
- [`docs/installer.md`](docs/installer.md) — the packaged Windows installer: auto-config engine, one-time `register-app`, and the SEA/Inno build.

## License

[MIT](LICENSE) © 2026 Tony Goodhew
