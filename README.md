# Claude Hotmail Connector

A **local** [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that lets Claude
**read email** and **read/write calendar events** in a **personal** Microsoft account
(`@hotmail.com` / `@outlook.com` / `@live.com`) via the Microsoft Graph API.

Claude's official Microsoft 365 connector only supports work/school (Microsoft Entra) accounts and
rejects personal Microsoft accounts. This project fills that gap as a **personal, single-user tool**
that runs entirely on your own machine.

> **Status:** the latest release (v0.1.0) is a Windows installer. The next release ships as a
> **Claude Desktop extension (`.mcpb`)** for Windows, macOS and Linux. The requirements are in
> [spec v2](docs/personal-outlook-mcp-connector-requirements.md), and the remaining work is tracked in
> the [Spec v2 epic (#66)](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/66).

## How it works — local stdio

The connector is a **local stdio MCP server** (spec v2 §4). Claude Desktop (as an extension) or Claude
Code / VS Code launches it as a subprocess and talks to it over stdin/stdout. There is **no public URL,
no tunnel and no hosting cost**, and no Claude ↔ server OAuth. The only sign-in is to **Microsoft**: a
native public-client OAuth 2.0 authorization-code + PKCE flow with a loopback redirect, opened in your
browser the first time a tool needs it. No client secret is involved.

```
Claude Desktop / Claude Code ──stdio──▶ this server ──HTTPS──▶ Microsoft Graph (personal account)
        (MCP client)                    (local subprocess)        graph.microsoft.com/v1.0
```

Because it runs locally, it works in **Claude Desktop** and **Claude Code / VS Code** on the same PC. It
does **not** work in claude.ai on the web or mobile; that would need a remote mode, which is a future
phase ([#58](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/58)).

## Tools

| Tool             | Purpose                                             | Graph permission      |
| ---------------- | --------------------------------------------------- | --------------------- |
| `whoami`         | Confirm the connected account                       | `User.Read`           |
| `list_messages`  | Search/list mail (find booking confirmations)       | `Mail.Read`           |
| `get_message`    | Read one email as text                              | `Mail.Read`           |
| `list_calendars` | List your calendars (own and shared)                | `Calendars.Read`      |
| `list_events`    | Read the calendar over a date range                 | `Calendars.Read`      |
| `create_event`   | Add a calendar event (timezone-correct, idempotent) | `Calendars.ReadWrite` |
| `update_event`   | Edit a calendar event (destructive: overwrites)     | `Calendars.ReadWrite` |
| `cancel_event`   | Delete an event, guarded (see below)                | `Calendars.ReadWrite` |
| `sign_out`       | Remove the saved sign-in from this computer         | (none)                |

Mail is **read-only** (no `Mail.Send`), and calendar writes never add attendees or send invitations.
There is **no hard-delete**. `cancel_event` moves an event to Outlook's Deleted Items, and only after
three checks:

- you've confirmed it and the event's exact title matches;
- the event has no attendees, since deleting a meeting could email cancellations;
- it isn't recurring.

`update_event` likewise refuses events with attendees, because changing a meeting emails its attendees.

Full tool contracts are in spec v2 §7.

## Step 1: register your own Microsoft Entra application

Every user needs their own registration; the connector doesn't ship a shared one, so each user owns
their own Microsoft access. Registering is free, but it needs a **Microsoft Entra tenant**, and a
personal account doesn't come with one. If you don't have a tenant,
[create a free one](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-create-new-tenant)
first (typically by signing up for a free Azure account).

> **Shortcut:** with the Azure CLI signed in to your tenant (`az login`),
> `pwsh ./scripts/register-app.ps1` does all of the steps below and prints the client ID.

The connector signs in as a **public/native client**. No client secret is ever used or stored.

1. Go to the [Entra admin center → App registrations](https://entra.microsoft.com/) →
   **New registration**.
2. **Name:** e.g. `Claude Hotmail Connector`.
3. **Supported account types:** choose **"Personal Microsoft accounts only"**. That's what the
   connector is for (`hotmail.com` / `outlook.com` / `live.com`), and it matches the default
   `MICROSOFT_TENANT=consumers`. Work/school accounts are served by Claude's official Microsoft 365
   connector. To allow them anyway, pick the multitenant + personal option instead and set
   `MICROSOFT_TENANT=common`.
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

## Step 2a: install in Claude Desktop (extension)

1. Download `hotmail-connector.mcpb` from
   [Releases](https://github.com/TGoodhew/Claude-Hotmail-Connector/releases) and verify it against
   `SHA256SUMS.txt`. (Until the next release, build it yourself: `npm run build:mcpb`, see
   [Development](#development).)
2. **Double-click it**, or drag it into **Claude Desktop → Settings → Extensions**, and choose
   **Install**.
3. Enter your **Application (client) ID** from step 1. You can change it later under **Settings →
   Extensions → Hotmail Connector**.
4. Ask Claude something that needs your mailbox or calendar. The first time, a Microsoft sign-in opens
   in your browser. Finish it within about 45 seconds and that request carries on; otherwise ask
   again once you've signed in.

> **Upgrading from the v0.1.0 Windows installer:** first uninstall **Claude Hotmail Connector** from
> Windows **Settings → Apps**. That removes its Claude Desktop entry; keeping both would show every
> tool twice. Then install the extension with **your own** client ID (v0.1.0's shared ID is being
> retired, [#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68)) and sign in again.

More detail, including how the bundle is built, is in [`docs/packaging.md`](docs/packaging.md).

## Step 2b: run from source (Claude Code / VS Code)

Prerequisites: **Node.js 24 LTS** is recommended; CI builds and releases on it. Node 20+ still
works (CI tests it, because the Claude Desktop extension declares that minimum), but Node 20
reached end-of-life in April 2026.

```bash
git clone https://github.com/TGoodhew/Claude-Hotmail-Connector.git
cd Claude-Hotmail-Connector
npm install
npm run build          # produces dist/index.js
```

Then point your client at `dist/index.js` and pass your client ID in the server's environment.

**Claude Code:**

```bash
claude mcp add --scope user --env MICROSOFT_CLIENT_ID=<your-client-id> \
  hotmail-connector -- node /path/to/Claude-Hotmail-Connector/dist/index.js
```

**VS Code:** use **Command Palette → "MCP: Add Server…" → Command (stdio)**, or copy
[`.vscode/mcp.json.example`](.vscode/mcp.json.example) to `.vscode/mcp.json` and edit the path and
client ID:

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

(Claude Desktop can also run a source checkout through `claude_desktop_config.json`, see
[`examples/claude_desktop_config.example.json`](examples/claude_desktop_config.example.json). The
extension is simpler.)

You don't need to sign in first: the first tool call opens the Microsoft sign-in in your browser. From
a source checkout you can also sign in, check or reset explicitly (with `MICROSOFT_CLIENT_ID` set, e.g.
in a `.env` in the checkout root, see [`.env.example`](.env.example). It's only ever read from there,
never from the folder you launch from):

```bash
node dist/index.js login     # browser sign-in, caches the token
node dist/index.js whoami    # prints your name and account, e.g. Jane Doe <jane@outlook.com>
node dist/index.js logout    # removes the saved sign-in (token cache and key)
```

### Configuration

| Variable                   | Default                                    | Notes                                                                                                                       |
| -------------------------- | ------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- |
| `MICROSOFT_CLIENT_ID`      | — (required)                               | Your Entra app's Application (client) ID. The extension sets it from its settings.                                          |
| `MICROSOFT_TENANT`         | `consumers`                                | Personal accounts only. `common` also allows work/school accounts (needs a multitenant registration).                       |
| `DEFAULT_TIMEZONE`         | This computer's time zone (UTC if unknown) | IANA zone used when a call doesn't specify one. An invalid name stops the server at startup.                                |
| `LOG_LEVEL`                | `info`                                     | `debug` / `info` / `warn` / `error` (logs go to stderr).                                                                    |
| `MICROSOFT_SCOPES`         | least-privilege set (spec §6)              | Narrowing only: drop scopes or use `Calendars.Read`; broader scopes are rejected. Tools needing a dropped scope are hidden. |
| `CLAUDE_HOTMAIL_CACHE_DIR` | per-OS profile dir                         | Where the encrypted token cache lives.                                                                                      |

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
  (`%LOCALAPPDATA%\claude-hotmail-connector` on Windows, `~/.claude-hotmail-connector` otherwise).
  On **Windows** the encryption key is protected with **DPAPI**, so only your Windows account on this
  PC can unlock it; copying the folder elsewhere isn't enough. On macOS/Linux the key is currently an
  owner-only file in the same folder; Keychain/libsecret support is tracked in
  [#83](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/83). Tokens are never committed,
  logged or returned to the client.
- **Logging:** structured logs go to **stderr only** (stdout is reserved for MCP JSON-RPC) and are
  scrubbed of tokens, `Authorization` headers, OAuth codes and secrets.
- **Email (and others' invites) are untrusted input:** they're written by strangers and may try to
  steer the model. The connector can't send mail: no `Mail.Send`, no attendees, and no changing or
  deleting meetings with attendees. Writes and deletes are marked destructive so Claude asks before
  running them, but only if your client's tool permissions ask. **Don't auto-approve the calendar
  write tools.** The full threat model is in
  [spec v2 §9.2](docs/personal-outlook-mcp-connector-requirements.md#92-threat-model). Review what
  Claude proposes before approving calendar changes.
- **Timezone-correct writes:** events are sent as local wall-clock time + an IANA time zone (never a
  UTC `Z` with a named zone), so they land at the intended local time.

Known security issues are tracked under the `area:security` label. To report a vulnerability, use
private reporting as described in [SECURITY.md](SECURITY.md); please don't open a public issue.

## Troubleshooting

- **`AADSTS50020` / "account does not exist in tenant" / personal account rejected:** the app
  registration doesn't accept personal accounts. Set its account type to **Personal Microsoft accounts
  only** (or multitenant + personal) per step 1.3.
- **`AADSTS9002331` / "configured for use by Microsoft Account users only":** a personal-only
  registration is being used with `MICROSOFT_TENANT=common`. Use the default, `consumers`.
- **A work/school account can't sign in:** expected with the default `consumers` tenant. See step 1.3
  for the multitenant option.
- **`AADSTS7000218` / "client_assertion or client_secret required":** the app is registered as a
  confidential/web client. Register it as a **public client** with **"Allow public client flows" =
  Yes** (step 1.6).
- **`redirect_uri` mismatch:** add **`http://localhost`** under the Mobile & desktop platform.
- **Event lands an hour off:** always pass a local wall-clock `dateTime` (e.g. `2026-07-20T10:00:00`)
  with an IANA `timeZone`, never a `Z`-suffixed value. The connector rejects `Z`/offset values for
  event times.
- **"No Microsoft app registration is configured":** no client ID is set. Enter it in **Claude Desktop
  → Settings → Extensions → Hotmail Connector**, or set `MICROSOFT_CLIENT_ID` in your client's server
  config (step 2b).
- **A tool says to finish signing in, or "sign-in is required":** finish the Microsoft sign-in in the
  browser window that opened, then ask Claude to retry. If no window appeared, check that a default web
  browser is set; from a source checkout you can also run `node dist/index.js login`. After a cancelled
  sign-in, the connector waits 30 seconds before opening another window.
- **Every tool appears twice:** you have both the extension and an older installer/manual config
  entry. Uninstall the v0.1.0 installer, or remove the `hotmail-connector` entry from
  `claude_desktop_config.json`.
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
npm run test:coverage   # the same, failing below the coverage thresholds (CI)
npm run test:live       # opt-in: real Graph, your account (after `node dist/index.js login`)
npm run build:mcpb      # Claude Desktop extension -> build/hotmail-connector.mcpb
```

Layout (tests are colocated as `src/**/*.test.ts`):

- `src/index.ts`: CLI and stdio entry (`login`, `logout`, `whoami`; default = server).
  `src/mcp.ts`: server assembly.
- `src/auth`: Microsoft OAuth (PKCE + loopback, on-demand sign-in) and the encrypted token cache
  (DPAPI-protected key on Windows).
- `src/graph`: Graph client (auth, retries, paging) plus the mail / calendar / user modules.
- `src/tools`: zod schemas and thin tool handlers.
- `src/util`: time zones, HTML-to-text, logging, errors.
- `scripts/`: `build-mcpb.mjs` (extension), `check-text.mjs`, `test-live.mjs`, `register-app.ps1`.

## Requirements, design & roadmap

- [`docs/personal-outlook-mcp-connector-requirements.md`](docs/personal-outlook-mcp-connector-requirements.md):
  the build spec, **v2** (tools, Graph, auth, security, timezone, packaging, decision log).
  Requirements not yet implemented link to their tracking issues.
- [`docs/packaging.md`](docs/packaging.md): the Claude Desktop extension (`.mcpb`): install, upgrade,
  build.
- [`docs/local-hosting-investigation-and-requirements.md`](docs/local-hosting-investigation-and-requirements.md):
  the original local-hosting investigation (superseded by spec v2; kept for the record).
- **Roadmap:** the [Spec v2 epic (#66)](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/66)
  orders all open work: spec decisions → security fixes → correctness → CI/tests → features.

## Support & contributing

This is a personal, open-source tool, **community-supported on a best-effort basis** with no SLA and
no hosted service. You run your own copy with your own Microsoft app registration. See
[SUPPORT.md](SUPPORT.md) for how to get help, [CONTRIBUTING.md](CONTRIBUTING.md) before opening a PR,
and [SECURITY.md](SECURITY.md) for private vulnerability reporting.

## License

[MIT](LICENSE) © 2026 Tony Goodhew
