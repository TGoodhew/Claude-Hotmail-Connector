# Security policy

## Reporting a vulnerability

Please report vulnerabilities **privately** through GitHub:
**[Report a vulnerability](https://github.com/TGoodhew/Claude-Hotmail-Connector/security/advisories/new)**
(Security tab → _Report a vulnerability_). Don't open a public issue, pull request or discussion for
a suspected vulnerability.

Include what you found, how to reproduce it, the affected version, and the impact you expect. Never
include real tokens, mail content or personal data.

This is a volunteer-maintained project (see [SUPPORT.md](SUPPORT.md)). Reports are handled on a
**best-effort basis with no guaranteed timelines**. Fixes are released through
[GitHub Releases](https://github.com/TGoodhew/Claude-Hotmail-Connector/releases).

## Supported versions

Only the **latest release** receives security fixes.

## Scope

In scope: this repository's code, build scripts, the Claude Desktop extension bundle and release artifacts. Examples include
token handling and storage, the OAuth sign-in flow, what the MCP tools expose to the model, and log
scrubbing.

Out of scope:

- vulnerabilities in Microsoft Entra / Graph, Claude, or third-party dependencies (report those
  upstream; we'll update affected dependencies);
- issues that require an already-compromised user account or PC;
- your own app registration's configuration.

Known security work is tracked publicly under the
[`area:security`](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues?q=is%3Aissue+label%3Aarea%3Asecurity)
label.
