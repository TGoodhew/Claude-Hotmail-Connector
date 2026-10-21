# Installer & packaging (Windows)

The connector ships as a **self-contained Windows package**: you can install it without Node and
without editing any JSON. You **do** need your own Microsoft Entra app registration (see below); the
connector deliberately doesn't ship a shared one (spec v2 §8.1, decision D3). This document covers the
user experience, the app registration, and how the artifacts are built.

Originally tracked in [#25](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/25); the
bring-your-own registration change is [#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68).

## What you get

Two downloads (both from the same single-exe build):

| Artifact                    | For         | Experience                                                    |
| --------------------------- | ----------- | ------------------------------------------------------------- |
| `HotmailConnectorSetup.exe` | most people | Guided Next → Finish installer, Start-menu entry, uninstaller |
| `hotmail-connector.exe`     | advanced    | Portable single exe, no install step                          |

## Before you install: your own app registration

The connector signs in to Microsoft through an Entra app registration that **you** own. It's a
**public/native** client (PKCE, no secret), so its **Application (client) ID** is not a secret.

1. **Get a Microsoft Entra tenant.** A personal Microsoft account (`hotmail` / `outlook` / `live`)
   doesn't have one by default. [Create a free tenant](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-create-new-tenant),
   typically by signing up for a free Azure account.
2. **Register the app**, either:
   - with one command (Azure CLI signed in to your tenant):
     ```pwsh
     az login
     pwsh ./scripts/register-app.ps1
     ```
   - or by hand in the portal: follow **"Register a Microsoft Entra application"** in the
     [README](../README.md).
3. Copy the printed **Application (client) ID**.

## Install journey

1. Download and run `HotmailConnectorSetup.exe`.
2. Windows shows an **"unknown publisher"** SmartScreen warning (the build is unsigned; see below).
   Click **More info → Run anyway**.
3. On **"Your Microsoft app registration"**, paste your Application (client) ID. When upgrading, you can
   leave it empty to keep the ID already saved on this PC.
4. The installer copies the exe and runs `hotmail-connector.exe setup --client-id <id>`. That command
   saves the ID to `%LOCALAPPDATA%\claude-hotmail-connector\settings.json` and **auto-configures Claude
   Desktop**. The finish page offers **"Sign in to your Microsoft account now"**.
5. **Fully quit Claude Desktop from the system tray** (closing the window only minimises it) and
   reopen it. The connector's tools are now available.

**Unattended install:** `HotmailConnectorSetup.exe /VERYSILENT /CLIENTID=<your-client-id>`, then run
`hotmail-connector.exe login` once.

### Why the SmartScreen warning (and why it's OK)

The package is **intentionally unsigned**. Since Microsoft changed SmartScreen reputation handling, a
standard code-signing certificate no longer removes the warning up front, so the cost isn't
worthwhile at this app's volume. The installer's welcome screen ([`installer/welcome.txt`](../installer/welcome.txt))
explains what the app is and why it's safe, and Releases publish SHA-256 checksums to verify the
download.

## Building the artifacts

> Run on **Windows**. Requires Node 20+ (24 LTS recommended, see #45) and
> [Inno Setup 6](https://jrsoftware.org/isinfo.php) (`iscc` on `PATH`) for the installer.

```pwsh
npm ci
npm run build:exe        # -> build\hotmail-connector.exe  (Node SEA single executable)
npm run build:installer  # -> dist-installer\HotmailConnectorSetup.exe
```

- `build:exe` ([`scripts/build-exe.mjs`](../scripts/build-exe.mjs)) bundles the app into a single CJS
  file, generates a Node **SEA** blob ([`sea-config.json`](../sea-config.json)), copies `node.exe`,
  and injects the blob with `postject`.
- `build:installer` compiles [`installer/hotmail-connector.iss`](../installer/hotmail-connector.iss).
- Releases are built and published by [`.github/workflows/release.yml`](../.github/workflows/release.yml).

### Smoke test the exe

```pwsh
build\hotmail-connector.exe setup --client-id <your-client-id>  # saves the id, configures Claude Desktop
build\hotmail-connector.exe login                              # one-time browser sign-in
build\hotmail-connector.exe whoami                             # prints the signed-in account
```

## The `setup` / `unsetup` commands

The exe (and `node dist/index.js`) expose these host-wiring commands, used by the installer:

- **`setup [--client-id <id>]`**
  - `--client-id` validates the ID (a GUID) and saves it to `settings.json` in the cache folder. That
    lets `login`, `whoami` and the Start-menu shortcut work without any environment variables.
  - Without the flag, setup uses `MICROSOFT_CLIENT_ID` or the saved ID. It refuses to run if neither
    is set.
  - It then detects the Claude Desktop configs: both the standalone `%APPDATA%\Claude\` path **and** the
    Microsoft Store/MSIX `…\Packages\Claude_*\LocalCache\Roaming\Claude\` path.
  - It merges our `mcpServers` entry in **without clobbering** existing preferences, writes valid
    JSON, and points `command` at the absolute exe/node path. Idempotent; backs up the original once.
- **`unsetup`**: removes only our entry (run on uninstall).

Neither command needs a working Microsoft configuration, so a bad environment can't block install or
uninstall. Claude Code is configured separately (see the main README).

## Status / remaining

- ✅ Single exe + Inno installer, published to GitHub Releases with checksums (v0.1.0).
- ✅ Bring-your-own registration: installer page, `setup --client-id`, saved settings (#68).
- ⏳ After the next release, the maintainer's old shared registration (used by v0.1.0) is retired
  (#68). v0.1.0 users need to create their own registration and run `setup --client-id` + `login`.
- ⏳ Uninstall cleanup of the token cache and an explicit `AppId` (#43); OS-protected token storage
  (#42); MCPB packaging evaluation (#63); macOS/Linux setup (#56).
