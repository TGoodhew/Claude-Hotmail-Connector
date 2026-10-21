# Packaging: Claude Desktop extension (.mcpb)

The connector ships to Claude Desktop as an **MCP Bundle** (`hotmail-connector.mcpb`), Claude Desktop's
one-click extension format, on Windows, macOS and Linux. Claude Desktop runs it with its **own bundled
Node.js**, shows a settings screen for it, and launches the server itself. There's no installer, no
separate runtime and no config file to edit. Claude Code and VS Code use the server from a source
checkout instead (see the [README](../README.md)).

Decided in [#63](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/63) (spec v2 decision
D13). It replaces the v0.1.0 Windows exe + Inno Setup installer.

## Before you install: your own app registration

The connector signs in to Microsoft through an Entra app registration that **you** own: a
**public/native** client using PKCE, with no secret. Its **Application (client) ID** isn't a secret.

1. **Get a Microsoft Entra tenant.** A personal Microsoft account (`hotmail` / `outlook` / `live`)
   doesn't have one by default.
   [Create a free tenant](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-create-new-tenant),
   typically by signing up for a free Azure account.
2. **Register the app**, either:
   - with one command (Azure CLI signed in to your tenant):
     ```pwsh
     az login
     pwsh ./scripts/register-app.ps1   # personal accounts only; -Audience Multitenant to widen
     ```
   - or by hand in the portal, following **"Register a Microsoft Entra application"** in the
     [README](../README.md).
3. Copy the printed **Application (client) ID**.

## Install

1. Download `hotmail-connector.mcpb` from
   [Releases](https://github.com/TGoodhew/Claude-Hotmail-Connector/releases) and verify it against
   `SHA256SUMS.txt`.
2. **Double-click it** (or drag it into **Claude Desktop → Settings → Extensions**) and choose
   **Install**.
3. Enter your **Application (client) ID** when asked. You can change it later under **Settings →
   Extensions → Hotmail Connector**.
4. Ask Claude something that needs your mailbox or calendar. The first time, a Microsoft sign-in opens
   in your browser. Finish it and that request carries on, or ask again if it took longer than about
   45 seconds.

Uninstall or disable it from **Settings → Extensions**.

### Upgrading from v0.1.0 (Windows installer)

1. Uninstall the old **Claude Hotmail Connector** from Windows **Settings → Apps**. Its uninstaller
   also removes its entry from Claude Desktop's config. Leaving both would run two servers and show
   every tool twice.
2. Install the `.mcpb` as above, with **your own** client ID (v0.1.0's shared ID is being retired,
   [#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68)).
3. Sign in again when prompted.

## Building

```bash
npm ci
npm run build:mcpb   # -> build/hotmail-connector.mcpb
```

[`scripts/build-mcpb.mjs`](../scripts/build-mcpb.mjs) does three things:
1. Bundles the server and its dependencies into one file (`server/index.cjs`, via esbuild), so the
   bundle needs no `node_modules`.
2. Writes `manifest.json` (MCPB manifest spec 0.3). The version comes from `package.json`, and a
   required `client_id` setting is mapped to `MICROSOFT_CLIENT_ID`.
3. Validates and packs it with the official CLI, `@anthropic-ai/mcpb` (pinned in `devDependencies`).

The result is about 220 kB. CI builds it on every PR, and [`release.yml`](../.github/workflows/release.yml)
publishes it with a checksum when a `v*` tag matching `package.json` is pushed.

Inspect a bundle with `npx mcpb info build/hotmail-connector.mcpb`, or unpack it with `npx mcpb unpack …`.

## Notes and known gaps

- **Node runtime:** Claude Desktop's bundled Node version isn't documented, so the bundle targets the
  same minimum as `package.json` `engines.node` (Node 20+), which the manifest declares in
  `compatibility.runtimes.node`.
- **Sign-out:** an extension has no command line. Removing the extension doesn't clear the cached
  sign-in in `%LOCALAPPDATA%\claude-hotmail-connector` (or `~/.claude-hotmail-connector`). Tracked in
  [#43](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/43).
- **Key protection:** Windows uses DPAPI (#42); macOS/Linux use a plain owner-only key file for now
  ([#83](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/83)).
- **Signing:** bundles are currently unsigned (`mcpb sign` is available if that changes).
