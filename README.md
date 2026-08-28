# SendToOneNote

Bring classic Outlook's "Send to OneNote" back — for new Outlook, including Gmail
and IMAP accounts where add-ins can't run.

Drag an email from new Outlook into a folder → pick a OneNote section → the email
becomes an editable OneNote page: subject as the title, From/To/Date block, full
HTML body with images embedded.

## Why a folder?

New Outlook doesn't allow add-ins on non-Microsoft accounts and blocks direct
drag-and-drop onto other apps. Dragging to a folder is the one path that works
for every account type. Pin the drop folder to Quick Access and it's one drag +
one click.

## How it saves

If desktop OneNote is installed, SendToOneNote writes the page straight into
its local cache — instant, no Microsoft sign-in, and OneNote's own sync
carries it to the cloud from there. Otherwise it falls back to Microsoft
Graph, which needs sign-in. `settings.json` has a `"Backend"` setting to
override the automatic choice: `"graph"` always uses the cloud path,
`"desktop"` always requires desktop OneNote; the default `"auto"` picks
desktop OneNote when it's present.

## Install

1. Download the latest release zip, unzip, run SendToOneNote.exe.
   (SmartScreen may warn because the exe is unsigned: More info → Run anyway.)
   Until the first release is published, build from source (see below).
2. If you have desktop OneNote, there is no sign-in — just pick the folder.
   Otherwise, sign in with the Microsoft account whose OneDrive holds your
   notebooks (work/school or personal).
3. Choose your drop folder. Done — drag emails in.

## Company (work/school) accounts

Your organization may require an admin to approve the app once
(Notes.ReadWrite). If sign-in is blocked: ask your admin to consent, or register
your own free app ID (docs/entra-app-registration.md) and put it in
%APPDATA%\SendToOneNote\settings.json as "ClientIdOverride".

## Privacy

Your email content goes from the local .eml file directly to Microsoft Graph
under your own sign-in. Nothing is sent anywhere else. No telemetry.

With desktop OneNote, nothing is sent anywhere by this app; OneNote's own
sync is the only network activity. Tracking pixels in emails are stripped
before saving, so senders don't get a read receipt from your machine.

## Fidelity notes

The desktop path renders through OneNote's own HTML importer — closer to
the classic "Send to OneNote" button. On the Graph path, OneNote's API
sanitizes HTML: scripts, forms, and complex CSS are removed, so heavily
designed newsletters simplify. Text, links, tables, and images survive.
Remote images are downloaded and embedded so pages outlive expiring links.

## Troubleshooting

- Drag-out from new Outlook sometimes needs the message clicked first
  (known Outlook regression); "… → Save as" always works.
- Failed saves land in the drop folder's Failed subfolder, each with a
  matching .error.txt explaining what went wrong; full logs in
  %APPDATA%\SendToOneNote\logs.
- Emails dropped into the folder while SendToOneNote is not running are not
  picked up automatically — start the app and drag the file out and back in
  (or re-save it).
- `"ImageDiagnostics": true` in settings.json writes
  `Diagnostics\<email>\images.csv` (plus the downloaded images) next to the
  drop folder for every save, showing why each image was embedded or
  dropped.
- `"Backend": "desktop"` failing at startup means desktop OneNote isn't
  installed — the Store app "OneNote for Windows 10" doesn't count; you
  need the desktop Office application.

## Building from source

dotnet build / dotnet test / dotnet run --project src/SendToOneNote. MIT license.
