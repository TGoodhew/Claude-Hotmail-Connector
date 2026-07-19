# Local Hosting — Investigation & Requirements (Personal Outlook.com MCP Connector)

**Companion to:** `personal-outlook-mcp-connector-requirements.md` (the "main spec")
**Document type:** Investigation + requirements
**Version:** 1.0 · **Date:** 2026-07-19 · **Owner:** Tony Goodhew

> **TL;DR** — Yes, you *can* run the connector locally. Whether "local" works depends entirely on **which Claude you connect from**:
> - **Claude Desktop app or VS Code on your machine → run it locally with no public URL** (stdio transport). This is the recommended local setup.
> - **Claude.ai in the browser / Cowork in the cloud → your local server must be exposed via a secure tunnel** (it can still *run* on your laptop, but Claude reaches it through a public HTTPS URL).
> A bare `http://localhost:…` URL pasted into the web "Add custom connector" box will **never** work — and Part A explains exactly why.

---

## Part A — Investigation: why "localhost" seems not to work

### A.1 The core reason
Claude's **custom connectors** (the ones you add under *Customize → Connectors*) are **remote MCP servers**. When you use Claude in the browser or in a Cowork **cloud** session, the component that opens the connection to your server is **Anthropic's cloud infrastructure, not your computer**. Anthropic's own requirement is explicit: the MCP server *"must be reachable over the public internet from Anthropic's IP ranges. Servers hosted on a private corporate network, behind a VPN, or blocked by a firewall won't connect."*

`http://localhost:3000` (a.k.a. `127.0.0.1`) means *"this same machine."* When Anthropic's server dials that URL, it resolves to **Anthropic's** loopback, not your laptop. So the connection either fails or hits the wrong machine. Two further blockers stack on top:
- **HTTPS is required** in production; a local `http://` endpoint (or a self-signed cert) won't satisfy it.
- The MCP spec itself says local Streamable-HTTP servers **should bind to `127.0.0.1` and validate the `Origin` header** to prevent DNS-rebinding attacks — i.e., local HTTP servers are designed to be reached by *local* clients, not exposed to the public internet.

**This is not a bug you can configure away.** It's the difference between a *local* client (which launches the server on your machine) and a *cloud* client (which must reach a public URL).

### A.2 The detail that makes "local" possible anyway: MCP has two transports
| Transport | How the client connects | Needs a public URL? | Who uses it |
|---|---|---|---|
| **stdio** | The client **launches the server as a local subprocess** and talks over stdin/stdout | **No** | Claude **Desktop** app, VS Code / Claude Code, MCP Inspector |
| **Streamable HTTP** | The client makes HTTP(S) calls to an **MCP endpoint URL** | **Yes** (for cloud clients) | Claude.ai web / Cowork cloud custom connectors, remote hosts |

Because **stdio needs no URL at all**, a connector that runs as a local subprocess of Claude Desktop or VS Code is genuinely "running locally" — no Cloudflare, no tunnel, no public IP.

### A.3 Why *this* Cowork session couldn't reach a local server
This Cowork task is running in **Anthropic's cloud sandbox**. Its shell and tools cannot reach `localhost` on your computer in either Cowork mode — the cloud sandbox and your machine are separate networks. So for **cloud Cowork specifically**, a local stdio server is invisible; you'd need either a tunnel (Part B, Mode B) or a hosted deployment (the main spec). A local stdio server shines with the **desktop app** and **VS Code**, which run the MCP client *on your machine*.

### A.4 Decision matrix — pick your path
| You want to use the connector from… | Recommended mode | Public URL / tunnel? |
|---|---|---|
| **Claude Desktop app** on your PC | **Mode A** — local stdio server | No |
| **VS Code / Claude Code** on your PC | **Mode A** — local stdio server | No |
| **Claude.ai in a browser** or **Cowork cloud** | **Mode B** — local HTTP + tunnel, *or* deploy to Cloudflare (main spec) | Yes |
| **Just developing/testing** | **Mode C** — MCP Inspector / VS Code at `localhost` | No |

**Conclusion of investigation:** Running locally is feasible and, for a single user on one machine, is arguably the *simplest* option (Mode A). The requirements for each local mode follow.

---

## Part B — Local hosting requirements by mode

> These build on the **main spec**. Everything there about the **tools** (`list_messages`, `get_message`, `list_events`, `create_event`, …), **Microsoft Graph** calls, **timezone handling**, and **security best practices** still applies. This document only changes **transport, hosting, and how the Microsoft OAuth is done**.

### Mode A — Local stdio server (recommended for local use)

**What it is:** A Node/TypeScript MCP server run as a subprocess by Claude Desktop or VS Code. No web server, no public URL, no tunnel.

**Key simplification — no downstream OAuth server.** With stdio, the *client trusts the local subprocess*, so you **do not** implement PRM/DCR/PKCE for the Claude↔server hop (all of §8.2 of the main spec drops away). You only need the **upstream** Microsoft authentication.

**Recommended upstream auth for a local app — native/public client + loopback:**
- Register the Entra app as a **"Mobile and desktop application"** (public client), **not** a web app.
- **No client secret** (native apps are public clients — nothing secret can be safely stored on a user's machine). Use **OAuth 2.0 authorization code + PKCE**.
- **Redirect URI:** a loopback address, e.g. `http://localhost:<port>` (Microsoft allows loopback redirects for native apps).
- **Token cache:** store the refresh token in the **OS keychain / secure credential store** (macOS Keychain, Windows Credential Manager, libsecret on Linux) or, at minimum, an encrypted file with `600` permissions — never plaintext, never in the repo.
- Scopes unchanged: `openid profile email offline_access User.Read Mail.Read Calendars.ReadWrite`, authority `common`.
- First run triggers a one-time browser sign-in (opens the system browser to Microsoft, captures the loopback redirect); subsequent runs use the cached refresh token silently.

**Transport requirements:**
- Implement the **stdio** transport (`StdioServerTransport` in `@modelcontextprotocol/sdk`).
- **Do not** write anything but valid MCP JSON-RPC to **stdout**; use **stderr** for logs.

**Client wiring:**
- **Claude Desktop:** register the server as a local MCP server — either packaged as a **Desktop Extension (`.mcpb`)** (the current one-click method) or via the developer config (`claude_desktop_config.json`) pointing at `node /path/to/dist/index.js` (or an `npx` command).
- **VS Code / Claude Code:** *MCP: Add Server… → (stdio) command*, pointing at the built entrypoint.

**Prerequisites:** Node 20+, the built server on disk, the Entra public-client registration, and Claude Desktop (latest) and/or VS Code.

**Security requirements (Mode A specifics):**
- Public client → **no secret on disk**; rely on PKCE + OS keychain for the refresh token.
- Bind nothing to the network (stdio only); if you also expose HTTP for testing, bind to `127.0.0.1` and validate `Origin`.
- Same logging hygiene as the main spec (never log tokens).

**Cost:** **$0.** Runs on your machine; no hosting, no tunnel. Only caveat: it's available **only while that machine is on and the client is running**, and only to clients on that machine.

**Limitation:** Not usable from Claude.ai in the browser or from a **cloud** Cowork session — those don't launch local subprocesses. Use Mode B for those.

---

### Mode B — Local HTTP server + secure tunnel (use local code from Claude.ai cloud)

**What it is:** Run the **Streamable HTTP** server (as in the main spec) on your laptop, then expose it with a tunnel that provides a **public HTTPS URL**. Paste that URL as a custom connector. The code runs locally; Anthropic reaches it through the tunnel.

**When to choose it:** You specifically want the connector available in **Claude.ai web / Cowork cloud** but prefer not to deploy to Cloudflare Workers yet (e.g., for iterative local development against the real Claude client).

**Requirements:**
- Full **downstream MCP OAuth** (PRM/DCR/PKCE/audience validation) from §8–§9 of the main spec **still applies** — the connection is now genuinely remote/public.
- The local server **MUST** enforce HTTPS (the tunnel terminates TLS and forwards to your local port), validate `Origin`, and require auth on every request.
- Tunnel options:
  | Tunnel | URL stability | Account | Cost |
  |---|---|---|---|
  | **Cloudflare Quick Tunnel** (`cloudflared tunnel --url http://localhost:8787`) | **Ephemeral** random `*.trycloudflare.com` each run | none | **Free** |
  | **Cloudflare Named Tunnel** | **Stable** (your domain) | Cloudflare account + domain | Free tier; domain ~$8–15/yr |
  | **ngrok** | 1 **free static domain**; otherwise random | ngrok account | **Free** tier (with limits); paid from ~$8–10/mo for more |

- Because Quick-Tunnel URLs change on each restart, you'd re-point the connector each time; a **named tunnel or ngrok static domain** gives a stable URL worth setting once.

**Security requirements (Mode B specifics):**
- Treat the tunnel URL as fully public: **auth is mandatory** (never expose unauthenticated tools).
- Prefer the upstream **web-app** Entra registration (client ID + **secret** stored as a local env secret, not committed) since this is a confidential-server pattern; or reuse the native pattern if the OAuth library supports it.
- Keep the machine patched; the tunnel is an inbound path to a process on your computer.

**Cost:** **$0** with Cloudflare Quick Tunnel or ngrok free tier; small optional costs for a custom domain. Same always-on caveat as Mode A — Claude can only reach it while your machine **and** the tunnel are running.

---

### Mode C — Local development & testing (not a connector)
- **MCP Inspector:** `npx @modelcontextprotocol/inspector` → point at the local server to browse tools, schemas, and exercise auth interactively.
- **VS Code MCP client:** add the server at `http://localhost:<port>` (HTTP) or as a stdio command; invoke tools from chat.
- Purpose: verify behavior before deploying (main spec) or wiring into Desktop (Mode A). Not intended as a durable "connector."

---

## Part C — Requirements common to all local modes
- **Reuse the main spec** for tool contracts, Graph endpoints, pagination/throttling, idempotency (`transactionId`), and **timezone correctness** (local wall-clock + IANA `Australia/Brisbane`).
- **Secrets & tokens:** never in source control; OS keychain or encrypted file (Mode A) / local env secrets (Mode B). `.gitignore` must cover `.env`, `.dev.vars`, token caches.
- **Least privilege:** `Mail.Read` + `Calendars.ReadWrite` only.
- **Logging:** structured, to stderr (Mode A) or a file (Mode B); never log tokens/headers.
- **Testing:** the same unit/integration tests from the main spec (§13), plus a transport-specific smoke test (stdio subprocess launch for Mode A; tunnel URL reachability for Mode B).

---

## Part D — Recommendation
- **If you'll use the connector from the Claude Desktop app or VS Code on your own computer:** choose **Mode A (local stdio)**. It's the simplest, cheapest ($0), and most private option, and it drops the entire downstream-OAuth machinery. Best default for a single user.
- **If you need it inside Claude.ai in the browser or Cowork cloud** (like this session): you need a public URL — either **Mode B (local + tunnel)** for a run-on-my-laptop setup, or the **Cloudflare Workers deployment** in the main spec for an always-on, hands-off option. For "set it and forget it," the hosted deployment wins; for "runs only when I'm at my desk," Mode B is fine.

A reasonable path: build **Mode A** first (fastest to a working local tool), and keep the code factored (per the main spec's `src/graph/*` separation) so the same tool logic can later be wrapped in the Streamable-HTTP + OAuth server for Mode B / hosted use with minimal change.

---

## Part E — References
- Build custom connectors (public-URL requirement): https://support.claude.com/en/articles/11503834-build-custom-connectors-via-remote-mcp-servers
- Get started with custom connectors: https://support.claude.com/en/articles/11175166-get-started-with-custom-connectors-using-remote-mcp
- Local MCP servers on Claude Desktop: https://support.claude.com/en/articles/10949351-getting-started-with-local-mcp-servers-on-claude-desktop
- MCP transports (stdio vs Streamable HTTP): https://modelcontextprotocol.io/docs/concepts/transports
- MCP TypeScript SDK (stdio + HTTP): https://github.com/modelcontextprotocol/typescript-sdk
- MCP Inspector: https://github.com/modelcontextprotocol/inspector
- Microsoft native/public client + loopback redirect (auth code + PKCE): https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-registration
- Microsoft auth code flow: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow
- Cloudflare Quick Tunnels (`trycloudflare`): https://developers.cloudflare.com/cloudflare-one/networks/connectors/cloudflare-tunnel/do-more-with-tunnels/trycloudflare/
- Share a local dev server via Cloudflare Tunnel: https://developers.cloudflare.com/workers/local-development/local-dev-tunnels/
- ngrok free plan limits: https://ngrok.com/docs/pricing-limits/free-plan-limits

---

*End of document.*
