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
- [ ] Fresh first-run on a machine with desktop OneNote shows no sign-in step
- [ ] Image-heavy newsletter saves within seconds and is visible in OneNote immediately (no "Page Not Yet Available")
- [ ] Saved page contains no tracking pixel (check page for stray 1x1 images)
- [ ] `"Backend": "graph"` forces the cloud path (tooltip says cloud; sign-in appears; save works)
- [ ] `"ImageDiagnostics": true` writes Diagnostics\<email>\images.csv with sensible decisions
- [ ] Picker shows local-only notebooks on the desktop path
