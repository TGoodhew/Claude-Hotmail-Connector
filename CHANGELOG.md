# Changelog

## Unreleased

### Added

- **PDF attachments** ([#156](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/156)).
  `get_attachment` now reads the text of PDF attachments, such as booking confirmations and tickets,
  from the first 30 pages of PDFs up to 5 MB.
  - Reading runs in a separate worker with time and memory limits, so a malicious PDF can't stall
    the connector.
  - Password-protected PDFs and scanned PDFs (no text layer) are refused with a clear message.
  - The extension now includes `THIRD_PARTY_NOTICES.txt`, listing the bundled open-source packages
    and their licenses, including pdf.js. It grows from about 0.3 MB to about 0.9 MB.

### Changed

- **Node.js 22.12 or later is now required**
  ([#148](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/148)). Node 20 reached
  end-of-life in April 2026. This only affects running from source: the Claude Desktop extension runs
  on Claude Desktop's own Node (24.22 in current versions). Tests now use vitest 5.

### Fixed

- **"Microsoft refused the saved sign-in."** When a saved sign-in no longer works (for example, after
  changing the client ID), the message now says so and explains how to sign in again, instead of
  "Fatal: … invalid_grant" ([#161](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/161)).
- **`cancel_event` and sign-outs elsewhere.** If you sign out in another place (e.g. `logout` in a
  terminal) while `cancel_event` is checking an event, the delete no longer goes through
  ([#143](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/143)).
- **Calendar (`.ics`) attachments** are read more strictly: contradictory date types, invalid
  statuses and oversized values are skipped rather than returned
  ([#157](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/157)).
- **Windows:** saving the sign-in no longer fails now and then when two copies of the connector touch
  it at once ([#142](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/142)).

## v0.2.1 — 2026-10-23

A bug-fix release. If you're on v0.2.0, install this over it the same way. If Claude Desktop asks for
your Application (client) ID again, enter the same one.

### Fixed

- **`find_free_time` and all-day events**
  ([#158](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/158)).
  - **The bug:** an all-day **Busy** event was shifted by your time zone's offset, so part of that day
    showed as free. In Pacific time it covered midnight to 5 PM, leaving 5 PM to midnight "free".
  - **The fix:** all-day events are now read as whole days in the time zone you ask in (by default your
    computer's), with daylight saving applied. Asking for free time "on Thursday in Berlin" treats an
    all-day busy event as all of Thursday in Berlin.
  - Events just outside the range can no longer change the status reported inside it.
- **`scripts/register-app.ps1`**
  ([#160](https://github.com/TGoodhew/Claude-Hotmail-Connector/pull/160)).
  - **The bug:** it created app registrations without their Microsoft Graph permissions listed. Sign-in
    still worked, because personal accounts approve permissions at sign-in.
  - **The fix:** the permissions are now set, and the script stops with an error if they're missing.
    If you made your registration with the v0.2.0 script, you can add them under **API permissions**
    in the Entra portal: `openid`, `profile`, `email`, `offline_access`, `User.Read`, `Mail.Read` and
    `Calendars.ReadWrite`. Doing so is optional.

## v0.2.0 — 2026-10-23

v0.2.0 is a rebuild against [spec v2](https://github.com/TGoodhew/Claude-Hotmail-Connector/blob/main/docs/personal-outlook-mcp-connector-requirements.md). The
connector now ships as a **Claude Desktop extension (`.mcpb`)** and uses **your own** Microsoft Entra
app registration. It also adds new mail and calendar tools and tightens security throughout.

### Upgrading from v0.1.0

1. Uninstall **Claude Hotmail Connector** from Windows **Settings → Apps**. This removes the old
   installer's Claude Desktop entry. If you keep both, every tool shows up twice.
2. Create your own app registration ([README, step 1](https://github.com/TGoodhew/Claude-Hotmail-Connector#step-1-register-your-own-microsoft-entra-application)).
   v0.1.0's shared client ID is being retired after this release
   ([#68](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/68)), and it will stop working.
3. Install `hotmail-connector.mcpb`: in Claude Desktop, go to **Settings → Extensions → Advanced
   settings → Install Extension…**, then enter your client ID.
4. Sign in again when Claude first needs your mailbox or calendar.

### Install and sign-in

- **Claude Desktop extension** (`.mcpb`) for Windows, macOS and Linux. It replaces the Windows exe,
  the installer and the `setup` / `unsetup` commands (#63).
- **Bring your own registration.** The bundled client ID is gone (#68), and the extension asks for
  yours. Personal Microsoft accounts only, by default (#62).
- **Sign-in on demand.** When a tool needs your account, the Microsoft sign-in opens in your browser
  (#69).
  - If it is still pending, asking again re-opens the same page, in case you closed its tab (#139).
  - Running Claude Desktop and Claude Code at the same time shares one sign-in and one sign-out (#95).
- **`sign_out`** removes the saved sign-in from this computer (#43).

### Mail (read-only)

- **Folders** (#53).
  - `list_messages` leaves out Junk Email and Deleted Items, and the folders beneath them, unless you
    ask for them.
  - It can search a single folder, and a message in Junk Email is flagged as such.
  - The new `list_mail_folders` lists your folders.
- **`get_attachment`** reads text, HTML and calendar (`.ics`) attachments (#54).
  - Calendar files come back as their events: title, times, place, details.
  - PDFs and other binary files are refused for now; PDF reading is
    [#156](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/156).
- `get_message` bodies are capped and readable in parts (#38). HTML is converted to text more
  faithfully, including tables (#32, #125). The content type is reported correctly (#35).

### Calendar

- **`find_free_time`** finds free slots in a date range (#52).
  - It can limit the search to daily hours, e.g. 09:00–17:00.
  - It is correct across daylight-saving changes.
  - Tentative events count as busy unless you say otherwise.
- **`get_event`** reads one event in full: notes, attendees and the recurrence pattern (#50).
- **`list_events`** pages through results, says when results were cut short, and can read any of your
  calendars (#49). The new `list_calendars` lists them.
- **`create_event`** supports all-day events, free/busy status and recurring events (#51).
  - All-day events default to Free, as in Outlook (#135).
  - Retrying a create doesn't make a duplicate (#48).
- **`update_event`** reports before and after values for what it changed (#71).
- **`cancel_event`** moves an event to Deleted Items, with guards (#55, #137):
  - you confirm the exact title;
  - there are no attendees;
  - it is not recurring (neither a series nor one of its occurrences);
  - the event is re-checked before every delete.
- Calendar writes never add attendees or send invitations (#61). Times you give are always read in
  the time zone you give (#33), and the default zone is your computer's (#34).

### Security

- **Untrusted content.** Mail and calendar content is marked as untrusted third-party text, so
  instructions hidden in an email are treated as data (#88, #60).
- **Token protection.**
  - The saved sign-in's key is protected with Windows DPAPI (#42).
  - `MICROSOFT_SCOPES` can only narrow access (#70).
- **Hardening.**
  - Paging cursors are opaque, and requests only go to Microsoft Graph (#30, #31).
  - Only the real sign-in redirect is accepted (#41).
  - `.env` is read only from the checkout (#41).
- **Error messages** say what to do next and never repeat Graph's own text (#35).
- **Release builds** carry signed build provenance (#46), alongside `SHA256SUMS.txt`.

### Known issues

- **No approval prompts.** Claude Desktop 3.1623.4 doesn't show approval prompts for extension
  tools, even when they're set to "Needs approval"
  ([#138](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/138); reported to Anthropic).
  The connector's own checks still apply: `cancel_event` keeps all the guards above. `update_event`
  refuses events with attendees and changes only the version it checked, but it doesn't ask you to
  confirm the title, so check what Claude plans to change before it does.
- **macOS and Linux are untested.** The extension is built for both but has only been tested by hand
  on Windows ([#98](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/98)).

### Thanks

Thanks to [@masyarakatai](https://github.com/masyarakatai), who contributed:

- the transaction ID fix (#67);
- daylight-saving test cases, used in #151;
- tool-description ideas, carried into #110.

## v0.1.0 — 2026-07-20

First release: a Windows installer and portable exe, with a shared app registration.
