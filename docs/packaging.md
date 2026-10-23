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
   - with one command from a source checkout. This needs the
     [Azure CLI](https://aka.ms/installazurecli) (on Windows:
     `winget install -e --id Microsoft.AzureCLI`, then open a new terminal):
     ```pwsh
     az login --allow-no-subscriptions   # a free tenant usually has no subscription
     pwsh ./scripts/register-app.ps1     # personal accounts only; -Audience Multitenant to widen
     ```
   - or by hand in the portal, following **"Step 1: register your own Microsoft Entra application"**
     in the [README](../README.md#step-1-register-your-own-microsoft-entra-application).
3. Copy the printed **Application (client) ID**.

## Install

1. Download `hotmail-connector.mcpb` from the
   [latest release](https://github.com/TGoodhew/Claude-Hotmail-Connector/releases/latest) and verify it against
   `SHA256SUMS.txt`. Releases also carry signed build provenance, so with the GitHub CLI you can
   check the bundle was built by this repository's release workflow:
   `gh attestation verify hotmail-connector.mcpb --repo TGoodhew/Claude-Hotmail-Connector`.
2. In Claude Desktop, open **Settings → Extensions → Advanced settings → Install Extension…**,
   choose the file, then **Install**. (Double-clicking it, or dragging it onto the Extensions page, may
   also work in some versions; dropping it on a chat only attaches it. Claude Desktop 3.1623.4 on
   Windows, the MSIX build, registers no file association for `.mcpb`.)
3. Enter your **Application (client) ID** when asked. You can change it later under **Settings →
   Extensions → Hotmail Connector**. Make sure the extension is switched on there. If you also run a
   source checkout, give it the same client ID. On Windows each keeps its own saved sign-in, because
   Claude Desktop is an MSIX app (see **Cache location** below); elsewhere they share one when they
   use the same cache folder.
4. Ask Claude something that needs your mailbox or calendar. The first time, a Microsoft sign-in opens
   in your browser. Finish it and that request carries on, or ask again if it took longer than about
   45 seconds.

Uninstall or disable it from **Settings → Extensions**. Removing the extension doesn't delete the
saved sign-in, so to leave nothing behind, first ask Claude to **sign out** (the `sign_out` tool).

### Updating to a newer release

Install the new `.mcpb` the same way, over the old one. If Claude Desktop asks for your client ID
again, enter the same one. What changed is in the [changelog](../CHANGELOG.md).

### Upgrading from v0.1.0 (Windows installer)

1. Uninstall the old **Claude Hotmail Connector** from Windows **Settings → Apps**. Its uninstaller
   also removes its entry from Claude Desktop's config. Leaving both would run two servers and show
   every tool twice.
2. Install the `.mcpb` as above, with **your own** client ID. v0.1.0's shared registration has been
   deleted, so v0.1.0 can no longer sign in
   ([#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68)).
3. Sign in again when prompted.

## Building

```bash
npm ci
npm run build:mcpb   # -> build/hotmail-connector.mcpb
```

[`scripts/build-mcpb.mjs`](../scripts/build-mcpb.mjs) does three things:
1. Bundles the server and its dependencies into one file (`server/index.cjs`, via esbuild), so the
   bundle needs no `node_modules`, and the PDF worker likewise (`server/pdf-worker.cjs`). It lists
   every bundled package and its license in `THIRD_PARTY_NOTICES.txt`, and copies in `LICENSE` and
   the icon.
2. Writes `manifest.json` (MCPB manifest spec 0.3). The version comes from `package.json`. The
   settings map to environment variables: `client_id` (required) to `MICROSOFT_CLIENT_ID`,
   **Allow sending email** to `ALLOW_SEND_MAIL`, **Allow saving email drafts** to
   `ALLOW_MAIL_DRAFTS`, and **Write a log file** to `LOG_FILE`.
3. Validates and packs it with the official CLI, `@anthropic-ai/mcpb` (pinned in `devDependencies`).

The result is about 1 MB, mostly pdf.js. CI builds it on every PR, and [`release.yml`](../.github/workflows/release.yml)
publishes it with a checksum when a `v*` tag matching `package.json` is pushed.

Inspect a bundle with `npx mcpb info build/hotmail-connector.mcpb`, or unpack it with `npx mcpb unpack …`.

## Notes and known gaps

- **Node runtime:** the extension runs on Claude Desktop's own bundled Node (24.22 in Claude Desktop
  3.6556.1; the version isn't documented, it was read from the app). The manifest declares the same
  minimum as `package.json` `engines.node`, Node 22.12+, in `compatibility.runtimes.node` (#148).
- **Sign-out:** an extension has no command line, so ask Claude to sign out (`sign_out`). That
  deletes the token cache and its key. Removing the extension alone doesn't (Claude Desktop runs no
  uninstall hook).
- **Cache location:** the extension asks for `%LOCALAPPDATA%\claude-hotmail-connector` (or
  `~/.claude-hotmail-connector`). Claude for Windows is an MSIX package, though, and Windows redirects
  files a packaged app creates under AppData to a per-package folder:
  `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Local\claude-hotmail-connector`. So the
  extension doesn't share its sign-in with a source checkout. `node dist/index.js logout` and deleting
  the unredirected folder don't sign Claude Desktop out; `sign_out` does. The two still share the
  named-pipe sign-in lock, which is named after the path and isn't redirected (#176).
- **Log file (#178):** Claude Desktop doesn't keep an extension's stderr, so the **Write a log file**
  setting (`LOG_FILE`) also writes the logs to `connector.log` in the cache folder (on the MSIX build,
  the redirected folder above). Lines are the same redacted ones as stderr, each prefixed with the
  time and process id. The file never exceeds 2 MB: when a line would take it over, the older half is
  cut off at a line boundary. `sign_out` leaves it in place.
- **Key protection:** Windows uses DPAPI (#42); macOS/Linux use a plain owner-only key file. OS
  keychain support there isn't planned; PRs welcome
  ([#83](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/83)).
- **Contents:** `server/index.cjs` (the server), `server/pdf-worker.cjs` (the PDF text extractor,
  started as a worker thread, #156), `manifest.json`, `icon.png`, `LICENSE`, and `THIRD_PARTY_NOTICES.txt`. The
  notices list every bundled package with its license; they're generated at build time from what
  esbuild actually bundled, plus pdf.js's Apache-2.0 license (`licenses/`), since unpdf bundles pdf.js
  without it.
- **Icon:** `assets/icon.svg` is the source, and `assets/icon.png` is the 512×512 PNG the bundle uses
  (the size MCPB recommends). The design is generic, with no Microsoft or Anthropic marks (#179).
  After changing the SVG, render the PNG again with Edge from a page that shows just the SVG at
  512×512, then commit both files:

  ```powershell
  & "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe" --headless=new --hide-scrollbars `
    --force-device-scale-factor=1 --default-background-color=00000000 --window-size=512,512 `
    --screenshot=assets\icon.png "file:///C:/path/to/icon.html"
  ```

  Here `icon.html` is `<body style="margin:0">` plus the SVG inline, with `svg{display:block}`.

- **Signing:** bundles are currently unsigned (`mcpb sign` is available if that changes).
