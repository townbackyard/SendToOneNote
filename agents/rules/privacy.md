# Privacy & repo-hygiene rules

This is a public, open-source repo. These rules keep it that way safely.

## Never committed, never quoted

- **`fixtures/local/`** contains the owner's real emails (personal and financial details). It is gitignored. Never commit it, never copy its contents into synthetic fixtures, tests, docs, issues, PRs, or commit messages — not even "harmless-looking" fragments like names, booking numbers, or amounts. Synthetic fixtures use invented `example.com`/`example.net` data only.
- No tokens, keys, or secrets exist in this project by design (public-client OAuth has none). If one ever appears in a diff, stop and flag it.

## Public by design (don't "fix")

- The **default client ID** in `MsalTokenProvider` is a public identifier for a public-client OAuth app. Shipping it in the repo is the standard OSS pattern, not a leak. `ClientIdOverride` in settings.json is the escape hatch for locked-down orgs.

## Runtime data-handling invariants

- **No telemetry, ever.** No analytics, crash reporting, or phone-home of any kind without an explicit owner decision (and then opt-in only). This is a README promise.
- Email content flows **only** from the local .eml to Microsoft Graph under the user's own token. No third-party endpoints. The only outbound HTTP besides Graph/login is downloading images referenced by the email being saved — initiated per-save, per-user-action.
- Scopes stay exactly `User.Read` + `Notes.ReadWrite`. Widening scopes is a spec change, not an implementation detail.
- Logs (`%APPDATA%\SendToOneNote\logs`) may contain email subjects and file names — that's acceptable on the user's own machine, but log content never leaves the machine and never goes into issues/PRs verbatim without the user's say-so.
