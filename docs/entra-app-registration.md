# Registering the Entra app (one-time, owner or BYO)

1. https://entra.microsoft.com → **Entra ID → App registrations** → New registration.
   (Menu naming varies slightly as Microsoft updates the portal; "Identity →
   Applications → App registrations" is the same place.)
2. Name: `SendToOneNote`. Supported account types: **Accounts in any organizational
   directory and personal Microsoft accounts** (shows as "All Microsoft account
   users" on the Overview page). Redirect URI: leave blank for now. Register.
3. On the app page, copy the **Application (client) ID** — this replaces
   `MsalTokenProvider.DefaultClientId` (owner) or goes into `settings.json` →
   `ClientIdOverride` (BYO users).
4. Add the redirect URIs and enable public client flows. The Authentication blade
   is being redesigned ("Authentication (Preview)"), so use whichever path matches
   what you see:
   - **UI path:** on Overview → Essentials, click **"Add a Redirect URI"** (or open
     Authentication and look for Add a platform / platform settings) → platform
     **Mobile and desktop applications** → check
     `https://login.microsoftonline.com/common/oauth2/nativeclient` → also add
     `ms-appx-web://microsoft.aad.brokerplugin/{client-id}` (required for the
     Windows WAM broker; substitute your actual client ID) → Save. Then set
     **Allow public client flows** = Yes on the same blade.
   - **Manifest path (works regardless of portal version):** Manage → **Manifest**,
     then set/merge and Save:

     ```json
     "isFallbackPublicClient": true,
     "publicClient": {
         "redirectUris": [
             "https://login.microsoftonline.com/common/oauth2/nativeclient",
             "ms-appx-web://microsoft.aad.brokerplugin/{client-id}",
             "http://localhost"
         ]
     }
     ```

     (`isFallbackPublicClient` is the "Allow public client flows" toggle;
     `http://localhost` is the loopback fallback if broker sign-in ever falls
     back to the system browser.)
5. API permissions → Add a permission → Microsoft Graph → Delegated →
   `Notes.ReadWrite` (User.Read is present by default). Do NOT grant admin consent for
   the whole org unless you intend to.
6. Branding & properties → set Publisher domain to your verified domain.
7. Publisher verification (recommended before public release): with a Partner Center
   account whose MPN ID is verified and whose domain matches the publisher domain,
   enter the MPN ID under Branding & properties → Publisher verification → Verify.
