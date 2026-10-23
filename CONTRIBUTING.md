# Contributing

Thanks for your interest! This is a small, volunteer-maintained project (see [SUPPORT.md](SUPPORT.md)),
so a little coordination goes a long way.

- **Comment on an issue before starting a PR.** Work is tracked in the
  [open issues](https://github.com/TGoodhew/Claude-Hotmail-Connector/issues), and areas of the code may
  be mid-change. Asking first avoids conflicting or duplicated work.
- **Read the spec.** [`docs/personal-outlook-mcp-connector-requirements.md`](docs/personal-outlook-mcp-connector-requirements.md)
  (v2) defines the tool contracts, security rules (least privilege, no attendees, no hard delete,
  email is untrusted input) and timezone handling. Changes that broaden access need a clear safety
  story.
- **Keep PRs focused**: one issue per PR, with tests, and `Closes #N` in the description.

## Before you push

```bash
npm ci
npm run typecheck
npm run lint
npm run format:check   # npm run format to fix
npm run check:text     # no raw control characters in tracked files
npm run build && npm run test:coverage
npm run build:mcpb     # the Claude Desktop extension builds and validates
```

CI runs the same checks. Never commit secrets, tokens, `.env` files, or real mail/calendar data in
test fixtures.

## Security

Report vulnerabilities privately; see [SECURITY.md](SECURITY.md).
