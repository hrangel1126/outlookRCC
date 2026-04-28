# RCC Mail — New Outlook Add-in Session

## Why this folder exists
The original RCC add-in (`RCCAddinNew/RCCAddinv1`) was a **VSTO add-in** (VB.NET, .NET 4.7.2).
It created a custom "RCC" ribbon tab in Classic Outlook with two buttons:
- "Enviar Correo" → WinForms dialog to compose email from shared mailbox
- "Configuracion" → WinForms dialog to manage shared mailboxes

**Problem:** New Outlook (and Outlook on the Web) does NOT support VSTO or custom ribbon tabs.
It only supports **Office Web Add-ins** — web pages hosted over HTTPS and wired via a manifest XML.

This folder (`outlooknew/`) is the transition to the new platform.

---

## Technology Stack

| Layer | Technology |
|---|---|
| Manifest | XML (`manifest.xml`) |
| UI | HTML + CSS (no frameworks) |
| Logic | Vanilla JavaScript |
| Auth | MSAL.js (popup flow) — acquireTokenSilent / acquireTokenPopup |
| Email sending | Microsoft Graph API (`/users/{mailbox}/sendMail`) |
| Backend | Vercel serverless functions (Node.js, `api/` folder) |
| Settings storage | localStorage (browser) |

---

## Current Architecture (as of April 28, 2026)

### Role split
- **Outlook add-in** — only used to obtain a Graph-scoped access token via MSAL
- **Vercel web page** (`/src/send.html`) — standalone browser page to compose and send email

### Token flow (step by step)
1. User opens add-in in Outlook → clicks **Obtener Token**
2. MSAL tries `acquireTokenSilent` first (no popup if token is cached)
3. If silent fails → MSAL opens a **Microsoft login popup**
4. User logs in with their Microsoft/M365 account
5. Popup redirects to `https://oulookrcc.vercel.app/src/auth-redirect.html` (same origin as add-in)
6. `auth-redirect.html` calls `handleRedirectPromise()` → popup closes automatically
7. MSAL returns an **access token already scoped for Graph** (`Mail.Send`, `Mail.Send.Shared`, `User.Read`)
8. Token stored in `localStorage` under key `rcc_office_token`
9. Taskpane shows: ✓ Token obtenido | user@domain | Expira: HH:MM
10. User opens `https://oulookrcc.vercel.app/src/send.html` in browser
11. Token auto-fills (same Vercel domain = same localStorage)
12. User fills form → clicks **Enviar** → POST `/api/send-email`
13. Vercel uses token directly as Graph bearer (no OBO needed)
14. Graph API sends the email from the shared mailbox

### Taskpane buttons (home panel — `src/taskpane.html`)
| Button | Action |
|---|---|
| 🔑 Obtener Token | MSAL login popup → stores Graph access token |
| ✅ Verificar Estado | Reads stored token, decodes JWT, shows valid/expired |
| ⚙ Configuracion | Navigate to settings.html |

---

## File Structure

```
outlooknew/
├── manifest.xml            Office Add-in manifest — ribbon wiring
├── vercel.json             Vercel build config — explicit static + API builds
├── package.json            npm scripts for local dev (start, setup, install-certs)
├── server.js               Local HTTPS server (port 3000) — local testing only, NOT deployed
├── index.html              Redirects to src/taskpane.html
├── SESSION.md              This file
├── INSTALL.md              Step-by-step install instructions
├── expectedwork.md         Architecture / flow documentation
├── api/
│   ├── send-email.js       POST — uses Graph token directly to send email (no OBO)
│   ├── test-token.js       POST — decodes JWT for debugging
│   ├── token.js            Helper: refresh token via refresh_token grant (unused currently)
│   └── auth/
│       ├── login.js        OAuth login redirect (unused currently)
│       └── callback.js     OAuth callback handler (unused currently)
├── src/
│   ├── styles.css          Shared CSS for all pages
│   ├── msal-config.js      MSAL config — clientId, authority, redirectUri (Vercel URL)
│   ├── auth-redirect.html  MSAL popup redirect page — must be same origin as add-in
│   ├── taskpane.html/js    Home panel — Obtener Token + Verificar Estado + Configuracion
│   ├── settings.html/js    Manage shared mailboxes (localStorage)
│   ├── send.html/js        Standalone send form — browser page, no Office.js needed
│   ├── compose.html/js     Old compose form — kept but not linked from taskpane
│   └── commands.html       Required manifest placeholder (no logic)
└── assets/
    ├── icon-16.png / icon-32.png / icon-64.png / icon-80.png / icon-128.png
```

---

## Deployed URLs

| Resource | URL |
|---|---|
| Vercel root | https://oulookrcc.vercel.app/ |
| Taskpane (add-in home) | https://oulookrcc.vercel.app/src/taskpane.html |
| Send form (standalone) | https://oulookrcc.vercel.app/src/send.html |
| MSAL redirect page | https://oulookrcc.vercel.app/src/auth-redirect.html |
| Send email API | https://oulookrcc.vercel.app/api/send-email |
| Test token API | https://oulookrcc.vercel.app/api/test-token |
| GitHub repo | https://github.com/hrangel1126/outlookRCC |

Vercel auto-deploys on every `git push` to `main`.

---

## Azure App Registration

| Field | Value |
|---|---|
| Client ID | `870de84d-3b21-449c-bf57-4cb3c76f9893` |
| Authority | `https://login.microsoftonline.com/common` |
| Redirect URI (SPA) | `https://oulookrcc.vercel.app/src/auth-redirect.html` |
| Scopes requested | `Mail.Send`, `Mail.Send.Shared`, `User.Read` |

---

## Vercel Environment Variables Required

| Variable | Value / Purpose |
|---|---|
| `API_KEY` | `rcc-api-key-2026` — must match value in `compose.js` and `send.js` |
| `CLIENT_ID` | Azure App Registration client ID (same as above) |
| `CLIENT_SECRET` | Azure App Registration client secret |
| `TENANT_ID` | Azure tenant ID (or `common`) |

> Note: `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` are no longer used by `send-email.js`
> (OBO was removed). They may still be needed by `auth/login.js` and `auth/callback.js`
> if those flows are ever activated.

---

## PENDING — Required Before Testing

### ⚠ Azure Portal step (one time, manual)
The MSAL redirect URI must be registered in Azure or the login popup will fail with
"redirect_uri_mismatch":

1. Go to **portal.azure.com** → **App registrations**
2. Find the app with client ID `870de84d-3b21-449c-bf57-4cb3c76f9893`
3. Click **Authentication** in the left menu
4. Under **Single-page application**, click **Add URI**
5. Add: `https://oulookrcc.vercel.app/src/auth-redirect.html`
6. Click **Save**

### ⚠ Clear Outlook add-in cache (after each redeployment)
Outlook's WebView2 caches add-in files aggressively. Uninstalling the manifest does NOT clear the file cache.

**To force a fresh load:**
1. Close Outlook completely
2. Delete everything inside: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
3. Reopen Outlook and re-sideload `manifest.xml`

Or inside the add-in panel: right-click → Inspect → Application → Clear Storage → Clear site data.

---

## How to Test End-to-End

1. Verify Vercel deployed: open `https://oulookrcc.vercel.app/src/send.html` — should show a form
2. Clear Office cache → re-sideload manifest → open add-in
3. Click **Obtener Token** → Microsoft login popup appears → log in → popup closes → status shows ✓
4. Open `https://oulookrcc.vercel.app/src/send.html` in the same browser Edge uses internally
5. Token should auto-fill in the Token field
6. Fill De (shared mailbox), Para, Asunto, Mensaje → click **Enviar**
7. Check that the email arrives at the destination

If token does not auto-fill: paste it manually from the "Token obtenido" display in the add-in.

---

## Known Issues / Notes

### Token not auto-filling in send.html
`localStorage` is only shared when `taskpane.html` and `send.html` run on the same origin AND
the same browser process. Outlook uses Edge WebView2 internally. If the user opens `send.html`
in a separate regular browser window, it may be a different storage partition.
**Workaround:** copy/paste the token manually.

### manifest.xml ribbon labels still say "Enviar Correo" / "Configuracion"
The ribbon buttons in the manifest open `compose.html` and `settings.html` respectively.
The new `taskpane.html` (with Obtener Token) is opened via the **Apps sidebar** in Outlook,
not the ribbon buttons. To update ribbon labels/targets, edit `manifest.xml` resources section.

### MSAL tokens expire in ~1 hour
After expiry, click **Obtener Token** again. If a cached account exists, it will silently refresh
without showing a popup.

---

## Key Technical Decisions & Lessons

| Decision | Reason |
|---|---|
| Switched from Office SSO to MSAL | `Office.auth.getAccessToken()` requires `<WebApplicationInfo>` in manifest + Azure SSO setup. MSAL popup works with the existing App Registration. Error 13000 = missing WebApplicationInfo. |
| Removed OBO from backend | MSAL returns a Graph-scoped token directly. OBO was needed for Office SSO bootstrap tokens only. |
| Explicit `builds` in `vercel.json` | Without it, Vercel detects `package.json` start script and runs `server.js` as a catch-all serverless function → 500 on every route. |
| `redirectUri` must match add-in origin | MSAL popup can only pass the token back to the opener if both are on the same domain. Old URI (GitHub Pages) was different from Vercel → token couldn't be returned. |

---
Saved: April 28, 2026
