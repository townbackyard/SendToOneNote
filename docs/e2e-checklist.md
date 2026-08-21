# Pre-release manual checklist

- [ ] Drag email from an M365 account in new Outlook → page correct (title, header, body)
- [ ] Drag email from a Gmail account in new Outlook → page correct
- [ ] Image-heavy HTML email → images embedded, page opens in OneNote client from toast
- [ ] Plain-text email → readable paragraphs / preserved columns
- [ ] Inline (cid:) image email → image embedded
- [ ] Esc in picker → .eml stays, nothing created
- [ ] Airplane mode → file lands in Failed with toast; drag back after reconnect succeeds
- [ ] Non-.eml file dropped → ignored, one log line
- [ ] Quit + relaunch → silent auth (no prompt), recents preserved
- [ ] Fresh Windows user / second machine → first-run flow works end to end
