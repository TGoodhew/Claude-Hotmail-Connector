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

## Tools (planned)

| Tool            | Purpose                                             | Graph permission      |
| --------------- | --------------------------------------------------- | --------------------- |
| `whoami`        | Confirm the connected account                       | `User.Read`           |
| `list_messages` | Search/list mail (find booking confirmations)       | `Mail.Read`           |
| `get_message`   | Read one email as text                              | `Mail.Read`           |
| `list_events`   | Read calendar over a date range                     | `Calendars.Read`      |
| `create_event`  | Add a calendar event (timezone-correct, idempotent) | `Calendars.ReadWrite` |
| `update_event`  | Edit a calendar event                               | `Calendars.ReadWrite` |

Mail is **read-only** (no `Mail.Send`). There is **no hard-delete** tool. `cancel_event` is deliberately
deferred as a gated/destructive action pending explicit approval.

## Requirements & design

The authoritative specifications live in [`docs/`](docs/):

- [`docs/personal-outlook-mcp-connector-requirements.md`](docs/personal-outlook-mcp-connector-requirements.md) — the main build spec (tools, Graph, security, timezone).
- [`docs/local-hosting-investigation-and-requirements.md`](docs/local-hosting-investigation-and-requirements.md) — why/how to host locally (Mode A stdio, chosen here).

## Status

Under active development. All work after this initial scaffold is tracked through GitHub issues
using a branch → work → commit → push → cleanup flow. See the
[issues](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues) for progress.

Setup, Microsoft Entra app registration, build, and client-wiring instructions are added by the
documentation issue once the implementation lands.

## License

[MIT](LICENSE) © 2026 Tony Goodhew
