# Installer & packaging (Windows)

This connector ships as a **self-contained Windows package** so a non-technical user can install it
without Node, without registering anything in Azure, and without editing any JSON. This document
covers the end-user experience, the maintainer's one-time setup, and how the artifacts are built.

Tracked in [#25](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/25).

## What the user gets

Two downloads (both from the same single-exe build):

| Artifact                    | For          | Experience                                                            |
| --------------------------- | ------------ | -------------------------------------------------------------------- |
| `HotmailConnectorSetup.exe` | most people  | Guided Next → Finish installer, Start-menu entry, uninstaller        |
| `hotmail-connector.exe`     | advanced     | Portable single exe, no install step                                 |

### End-user journey (zero portal, zero config)

1. Download and run the installer.
2. Windows shows an **"unknown publisher"** SmartScreen warning (the build is unsigned — see below).
   Click **More info → Run anyway**.
3. The installer copies the exe, **auto-configures Claude Desktop** (`hotmail-connector.exe setup`),
   and offers **"Sign in to your Microsoft account now"** on the finish page.
4. **Fully quit Claude Desktop from the system tray** (closing the window only minimises it) and
   reopen it. The connector's tools are now available.

No `MICROSOFT_CLIENT_ID`, no `.env`, no Azure portal: the build embeds a **shared public-client id**
(`DEFAULT_CLIENT_ID` in [`src/config.ts`](../src/config.ts)); the user only signs in.

### Why the SmartScreen warning (and why it's OK)

The package is **intentionally unsigned**. Since Microsoft changed SmartScreen reputation handling, a
standard code-signing certificate no longer removes the warning up front, so the cost isn't
worthwhile at this app's volume. The installer's welcome screen ([`installer/welcome.txt`](../installer/welcome.txt))
explains what the app is and why it's safe, and Releases publish checksums to verify the download.

## Maintainer: one-time app registration

The shared registration is created **once**, by the maintainer, in a free tenant they own — end users
never do this (a personal Microsoft account has no tenant to register an app in).

```pwsh
az login
pwsh ./scripts/register-app.ps1
```

This creates a **multitenant, public/native** client (PKCE, no secret) with the least-privilege
delegated scopes and prints the **Application (client) id**. Put that id in
[`src/config.ts`](../src/config.ts) → `DEFAULT_CLIENT_ID`. Because it's a public client, the id is not
a secret and is safe to embed and distribute.

Recommended before wide distribution: complete **Microsoft Publisher Verification** and set
consent-screen branding (logo, privacy/ToS URLs) so users don't see the "unverified app" prompt.
(A manual portal walkthrough is the fallback if you'd rather not use `az`.)

## Building the artifacts

> Run on **Windows**. Requires Node 20+ and [Inno Setup 6](https://jrsoftware.org/isinfo.php)
> (`iscc` on `PATH`) for the installer.

```pwsh
npm ci
npm run build:exe        # -> build\hotmail-connector.exe  (Node SEA single executable)
npm run build:installer  # -> dist-installer\HotmailConnectorSetup.exe
```

- `build:exe` ([`scripts/build-exe.mjs`](../scripts/build-exe.mjs)) bundles the app to a single CJS
  file, generates a Node **SEA** blob ([`sea-config.json`](../sea-config.json)), copies `node.exe`,
  and injects the blob with `postject`.
- `build:installer` compiles [`installer/hotmail-connector.iss`](../installer/hotmail-connector.iss).

### Smoke test the exe

```pwsh
build\hotmail-connector.exe whoami     # after signing in
build\hotmail-connector.exe setup      # configures Claude Desktop (idempotent, backs up first)
```

## The `setup` / `unsetup` commands

The exe (and `node dist/index.js`) expose these host-wiring commands, used by the installer:

- **`setup`** — detects Claude Desktop configs (both the standalone `%APPDATA%\Claude\` path **and**
  the Microsoft Store/MSIX `…\Packages\Claude_*\LocalCache\Roaming\Claude\` path) and merges our
  `mcpServers` entry in **without clobbering** existing preferences, writing valid JSON and pointing
  `command` at the absolute exe/node path. Idempotent; backs up the original once.
- **`unsetup`** — removes only our entry (run on uninstall).

Claude Code is configured separately via a project `.mcp.json` (see the main README).

## Status / remaining

- ✅ Auto-config engine, `setup`/`unsetup`, embedded client id, first-run wiring — implemented and
  unit-tested; `setup` verified against a real MSIX Claude Desktop config.
- ⏳ **Needs a Windows build run to finish #25:** produce `hotmail-connector.exe` (`build:exe`) and
  `HotmailConnectorSetup.exe` (`build:installer`), then smoke-test the installer on a clean machine
  and attach both to a GitHub Release with checksums. macOS/Linux packaging is out of scope for v1.
