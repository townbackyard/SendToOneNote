# SendToOneNote — Claude Code Configuration

@AGENTS.md

## Git Policy

- Commits and pushes that are part of an **approved plan task** (spec/plan/issue work the owner has green-lit) are fine — state what's being committed.
- Anything outside an approved plan (new files, restructures, policy changes like this file) needs explicit approval before commit.
- Tags and releases are owner decisions — never tag without an explicit go-ahead.
- Never use `--no-verify`, `--no-gpg-sign`, force-push, or history rewrites unless explicitly asked.

## Owner-Machine Policy

Changes to the owner's machine or accounts are **always performed manually by the owner** — propose exact steps, never execute:

- Outlook configuration, drag-and-drop tests, drop-folder creation on the real machine.
- Entra portal work (app registration, publisher verification) — see `docs/entra-app-registration.md` and issue #14.
- Installing the app, Startup shortcuts, signing in to Microsoft accounts.

## Fixtures

- `fixtures/local/` holds the owner's real emails (personal/financial data). **Never commit, quote, or copy their contents** into code, tests, docs, issues, or commit messages. Synthetic fixtures in `fixtures/synthetic/` exist precisely so real data never needs to travel.

## Known deliberate choices (don't re-flag)

- Unsigned exe releases (SmartScreen warning documented in README) — code signing is a cost decision deferred until traction.
- The shipped default client ID is public by design (public-client OAuth); `ClientIdOverride` exists for locked-down orgs.
- v1 shows a Quick Access *tip* instead of auto-pinning, and the picker has no background-refresh indicator — both tracked as issue #20, not bugs.
