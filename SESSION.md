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

| Layer | Technology | Replaces |
|---|---|---|
| Manifest | XML (`manifest.xml`) | VSTO project registration / registry |
| UI | HTML + CSS | WinForms (SettingsForm.vb, ComposeEmailForm.vb) |
| Logic | Vanilla JavaScript | VB.NET code in forms |
| Outlook API | Office.js (CDN, no install) | Microsoft.Office.Interop.Outlook |
| Settings storage | localStorage (browser) | `%APPDATA%\RCCAddIn\settings.txt` |
| Backend | Vercel serverless functions (Node.js) | Not needed in VSTO |
| Email sending | Microsoft Graph API (OBO flow) | Microsoft.Office.Interop.Outlook |

No frameworks — plain HTML/JS. Easy to read and maintain.

---

## Current Architecture (as of April 27, 2026)

### Role split
- **Outlook add-in** — only used to obtain the Office SSO token
- **Vercel web page** (`/src/send.html`) — standalone page to compose and send email using that token

### Token flow
1. User opens add-in in Outlook → clicks **Obtener Token**
2. `Office.auth.getAccessToken()` borrows the existing Outlook session (no popup)
3. Token stored in `localStorage` under `rcc_office_token`
4. User opens `https://oulookrcc.vercel.app/src/send.html` in a browser
5. Token auto-fills (same Vercel domain → same localStorage origin)
6. User fills form → clicks **Enviar** → POST `/api/send-email`
7. Vercel exchanges token via OBO flow → sends via Microsoft Graph API

### Taskpane buttons (home panel)
| Button | Action |
|---|---|
| 🔑 Obtener Token | Gets Office SSO token, stores it, shows user + expiry |
| ✅ Verificar Estado | Decodes stored token, shows if valid or expired |
| ⚙ Configuracion | Navigate to settings.html |

---

## File Structure

```
outlooknew/
├── manifest.xml          Office Add-in manifest — ribbon wiring
├── vercel.json           Vercel build config — static files + API functions
├── package.json          npm scripts: start (local dev), setup, install-certs
├── server.js             Local HTTPS server (port 3000) — local testing only
├── setup.js              Generates placeholder PNG icons in assets/
├── index.html            Redirects to src/taskpane.html
├── SESSION.md            This file
├── INSTALL.md            Step-by-step install instructions
├── expectedwork.md       Architecture / flow documentation
├── api/
│   ├── send-email.js     POST — receives token + email data, sends via Graph API
│   ├── test-token.js     POST — decodes Office JWT for debugging
│   ├── token.js          Helper: refresh token via refresh_token grant
│   └── auth/
│       ├── login.js      OAuth login redirect
│       └── callback.js   OAuth callback handler
├── src/
│   ├── styles.css        Shared CSS for all pages
│   ├── taskpane.html/js  Home panel (Obtener Token + Verificar Estado + Configuracion)
│   ├── settings.html/js  Manage mailboxes form
│   ├── send.html/js      Standalone send form — works outside Outlook, hosted on Vercel
│   ├── msal-config.js    MSAL config (kept but not used by current flow)
│   ├── compose.html/js   Old compose form (kept, not linked from taskpane anymore)
│   ├── auth-redirect.html MSAL redirect page (kept)
│   └── commands.html     Required manifest placeholder (no logic)
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    ├── icon-80.png
    └── icon-128.png
```

---

## Deployed URLs

| Resource | URL |
|---|---|
| Vercel home | https://oulookrcc.vercel.app/ |
| Taskpane (add-in) | https://oulookrcc.vercel.app/src/taskpane.html |
| Send form (standalone) | https://oulookrcc.vercel.app/src/send.html |
| Send email API | https://oulookrcc.vercel.app/api/send-email |
| Test token API | https://oulookrcc.vercel.app/api/test-token |
| GitHub repo | https://github.com/hrangel1126/outlookRCC |

---

## Vercel Environment Variables Required

| Variable | Purpose |
|---|---|
| `API_KEY` | Must match `rcc-api-key-2026` in client code |
| `CLIENT_ID` | Azure App Registration client ID |
| `CLIENT_SECRET` | Azure App Registration client secret |
| `TENANT_ID` | Azure tenant ID (or "common") |

---

## Known Issues / Next Steps

### Add-in cache in Outlook
After redeploying, Outlook's WebView2 cache may serve old files.
To clear: close Outlook → delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` → reopen.

### Token auto-fill on send.html
`localStorage` is shared between `taskpane.html` and `send.html` only if both are on the same
origin (`oulookrcc.vercel.app`). If the user opens `send.html` in a different browser than
Outlook uses internally (Edge WebView2), they will need to paste the token manually.

### localStorage does not roam across devices
If cross-device sync is needed, switch to `Office.context.roamingSettings` (Exchange-synced, max 32KB).

### MSAL config still present but unused
`msal-config.js` and the MSAL CDN script were removed from `taskpane.html`.
`taskpane.js` no longer uses MSAL — all auth goes through Office SSO.
`compose.html` and `auth-redirect.html` still reference MSAL but are no longer linked from the taskpane.

---

## Lessons from the VSTO Project

- Use HKEY_CURRENT_USER — no admin rights required
- Duplicate ribbon files break the project (don't replicate that pattern here)
- .NET 4.7.2 / AnyCPU constraints do NOT apply to this web add-in
- Settings in localStorage do NOT roam across devices

---
Saved: April 27, 2026
