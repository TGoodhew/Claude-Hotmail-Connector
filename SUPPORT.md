# Support

Claude Hotmail Connector is a personal, open-source tool shared **as is** under the [MIT License](LICENSE).
It is **community-supported on a best-effort basis**: there is no SLA, no guaranteed response time, and
no hosted service behind it.

## You run your own copy

- The connector runs **locally** on your PC; nothing is hosted for you.
- You sign in through **your own** Microsoft Entra app registration
  ([README step 1](README.md#1-register-a-microsoft-entra-application)). You own and are responsible for
  that registration, your tenant, and the access you grant to it. The maintainer has no access to your
  account, mail, calendar or tokens, and can't recover or reset them.

## Getting help

1. Check [Troubleshooting](README.md#troubleshooting) in the README and
   [`docs/packaging.md`](docs/packaging.md) (the Claude Desktop extension).
2. Search [existing issues](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues?q=is%3Aissue).
3. If it's a bug, open an issue with the **Bug report** template. Include the version, your OS, the
   client (Claude Desktop / Claude Code), and a `LOG_LEVEL=debug` stderr excerpt. **Redact** email
   addresses, message content and anything that looks like a token first.

Feature ideas are welcome via the **Feature request** template. The roadmap lives in the
[Spec v2 epic (#66)](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues/66).

## Security issues

Don't open a public issue for a vulnerability. See [SECURITY.md](SECURITY.md).
