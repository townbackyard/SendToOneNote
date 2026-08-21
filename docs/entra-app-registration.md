# Registering the Entra app (one-time, owner or BYO)

1. https://entra.microsoft.com → Identity → Applications → App registrations → New registration.
2. Name: `SendToOneNote`. Supported account types: **Accounts in any organizational
   directory and personal Microsoft accounts**. Redirect URI: leave blank for now. Register.
3. On the app page, copy the **Application (client) ID** — this replaces
   `MsalTokenProvider.DefaultClientId` (owner) or goes into `settings.json` →
   `ClientIdOverride` (BYO users).
4. Authentication → Add a platform → **Mobile and desktop applications** → check
   `https://login.microsoftonline.com/common/oauth2/nativeclient` → also add redirect URI
   `ms-appx-web://microsoft.aad.brokerplugin/{client-id}` (required for WAM broker) → Save.
   Set **Allow public client flows** = Yes.
5. API permissions → Add a permission → Microsoft Graph → Delegated →
   `Notes.ReadWrite` (User.Read is present by default). Do NOT grant admin consent for
   the whole org unless you intend to.
6. Branding & properties → set Publisher domain to your verified domain.
7. Publisher verification (recommended before public release): with a Partner Center
   account whose MPN ID is verified and whose domain matches the publisher domain,
   enter the MPN ID under Branding & properties → Publisher verification → Verify.
