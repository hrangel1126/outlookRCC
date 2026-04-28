# RCC Mail — New Outlook Add-in Session

## What this project is
A replacement for the old VSTO/VB.NET Outlook add-in.
New Outlook does not support VSTO — it only supports Office Web Add-ins (HTML + JS).
This add-in lets users send emails from shared mailboxes via Microsoft Graph API.

---

## Final Architecture (April 28, 2026)

### Two separate hosts — no Azure changes needed
| Host | What it serves | URL |
|---|---|---|
| GitHub Pages | All HTML/JS/CSS files (the add-in UI) | `https://hrangel1126.github.io/outlookRCC/` |
| Vercel | API serverless functions only | `https://oulookrcc.vercel.app/api/` |

**Why two hosts?**
- MSAL needs the redirect URI (where the login popup lands) to be on the **same domain** as the page that opened the popup
- The only registered redirect URI is `https://hrangel1126.github.io/outlookRCC/src/auth-redirect.html` — already set up, no Azure changes needed
- So all add-in HTML pages must also be on `hrangel1126.github.io`
- Vercel still hosts the API endpoints and adds CORS headers so GitHub Pages can call them

### Authentication — MSAL popup (no Azure App Registration changes needed)
- Uses existing Azure App Registration: `clientId: 870de84d-3b21-449c-bf57-4cb3c76f9893`
- Registered redirect URI: `https://hrangel1126.github.io/outlookRCC/src/auth-redirect.html`
- MSAL requests scopes: `Mail.Send`, `Mail.Send.Shared`, `User.Read`
- Token is already a Graph access token — no OBO exchange needed on the backend

### Full flow (step by step)
1. User opens the add-in in Outlook (via ribbon "Obtener Token" button or Apps sidebar)
2. Add-in loads `taskpane.html` from GitHub Pages
3. User clicks **🔑 Obtener Token**
4. MSAL tries silent token first (cached) — if it works, no popup
5. If no cached token → Microsoft login popup appears
6. User logs in with their Microsoft/M365 account
7. Popup redirects to `auth-redirect.html` (GitHub Pages — same origin)
8. `auth-redirect.html` calls `handleRedirectPromise()` → popup closes automatically
9. Token stored in `localStorage` under `rcc_office_token` (GitHub Pages origin)
10. Taskpane shows: ✓ Token obtenido | user@domain.com | Expira: HH:MM
11. User opens `https://hrangel1126.github.io/outlookRCC/src/send.html` in browser
12. Token auto-fills (same GitHub Pages origin = same localStorage)
13. User fills: De (shared mailbox), Para, Asunto, Mensaje → clicks **Enviar**
14. `send.js` POSTs to `https://oulookrcc.vercel.app/api/send-email` with token + data
15. Vercel uses the Graph token directly to send via Microsoft Graph API
16. Email sent from the shared mailbox ✓

---

## Ribbon Buttons (manifest.xml)

| Ribbon button | Opens | Label |
|---|---|---|
| Obtener Token | `taskpane.html` (GitHub Pages) | Obtener Token |
| Configuracion | `settings.html` (GitHub Pages) | Configuracion |

The **Apps sidebar** in Outlook also opens `taskpane.html`.

---

## File Structure

```
outlooknew/                         ← repo root (git)
├── manifest.xml                    ← sideload this in Outlook
├── vercel.json                     ← explicit builds: api/*.js=Node, src/**=static
├── package.json                    ← npm scripts for local dev only
├── server.js                       ← local HTTPS dev server — NOT deployed
├── index.html                      ← redirects to src/taskpane.html
├── api/
│   ├── send-email.js               ← POST: Graph token → send email (CORS enabled)
│   ├── test-token.js               ← POST: decode JWT for debugging (CORS enabled)
│   ├── token.js                    ← refresh token helper (unused currently)
│   └── auth/login.js, callback.js  ← OAuth helpers (unused currently)
├── src/
│   ├── styles.css                  ← shared styles
│   ├── msal-config.js              ← clientId, authority, redirectUri (GitHub Pages)
│   ├── auth-redirect.html          ← MSAL popup redirect handler (must stay on GitHub Pages)
│   ├── taskpane.html / .js         ← add-in home: Obtener Token + Verificar Estado + Config
│   ├── settings.html / .js         ← manage shared mailboxes (localStorage)
│   ├── send.html / .js             ← standalone send form (browser, no Office.js needed)
│   └── commands.html               ← manifest placeholder (no logic)
└── assets/
    └── icon-16/32/64/80/128.png
```

---

## All URLs

| Resource | URL |
|---|---|
| Add-in home (taskpane) | https://hrangel1126.github.io/outlookRCC/src/taskpane.html |
| MSAL redirect handler | https://hrangel1126.github.io/outlookRCC/src/auth-redirect.html |
| Send form | https://hrangel1126.github.io/outlookRCC/src/send.html |
| Settings | https://hrangel1126.github.io/outlookRCC/src/settings.html |
| Send email API | https://oulookrcc.vercel.app/api/send-email |
| Test token API | https://oulookrcc.vercel.app/api/test-token |
| GitHub repo | https://github.com/hrangel1126/outlookRCC |

Vercel auto-deploys on `git push` to `main`.
GitHub Pages serves files from the `main` branch root automatically.

---

## Azure App Registration (existing — no changes needed)

| Field | Value |
|---|---|
| Client ID | `870de84d-3b21-449c-bf57-4cb3c76f9893` |
| Authority | `https://login.microsoftonline.com/common` |
| Redirect URI (SPA) | `https://hrangel1126.github.io/outlookRCC/src/auth-redirect.html` |
| Scopes | `Mail.Send`, `Mail.Send.Shared`, `User.Read` |

---

## Vercel Environment Variables

| Variable | Value |
|---|---|
| `API_KEY` | `rcc-api-key-2026` (must match value in `send.js`) |
| `CLIENT_ID` | Azure App client ID (kept for future use) |
| `CLIENT_SECRET` | Azure App client secret (kept for future use) |
| `TENANT_ID` | Azure tenant ID (kept for future use) |

---

## How to Test End-to-End

1. **Verify GitHub Pages is live:**
   Open `https://hrangel1126.github.io/outlookRCC/src/taskpane.html` in a browser — should show the RCC Mail panel.

2. **Verify Vercel API is live:**
   The Vercel deployment should show static files at `/src/send.html` without 500 errors.

3. **Clear Outlook add-in cache** (every time after manifest changes):
   - Close Outlook
   - Delete everything in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Reopen Outlook

4. **Re-sideload the manifest:**
   - Outlook → Apps → My add-ins → Add from file → select `manifest.xml`

5. **Get the token:**
   - Click "Obtener Token" ribbon button (or find it in the Apps sidebar)
   - Microsoft login popup appears → log in → popup closes
   - Status shows ✓ Token obtenido | your@email | Expira: HH:MM

6. **Send a test email:**
   - Open `https://hrangel1126.github.io/outlookRCC/src/send.html` in Edge (same browser Outlook uses internally)
   - Token should auto-fill
   - Fill De (shared mailbox address), Para, Asunto, Mensaje
   - Click Enviar → check that email arrives

7. **If token does not auto-fill:**
   Copy the token text shown in the add-in and paste it manually into the Token field on send.html.

---

## Key Technical Lessons (do not repeat these mistakes)

| Problem | Root cause | Fix |
|---|---|---|
| Error 13000 on `Office.auth.getAccessToken()` | Manifest missing `<WebApplicationInfo>` + Azure SSO not configured | Switch to MSAL popup instead |
| 500 on all Vercel routes | `package.json` has `start: node server.js` → Vercel runs server.js as catch-all serverless function → crashes on missing SSL certs | Use explicit `builds` array in `vercel.json` |
| MSAL popup token not returned to opener | `taskpane.html` on Vercel + `auth-redirect.html` on GitHub Pages = different origins → MSAL storage-based messaging fails cross-origin | Host all add-in HTML on GitHub Pages (same origin as registered redirect URI) |
| "redirectUri_mismatch" from Azure | redirectUri in code must exactly match what is registered in Azure App Registration | Use `https://hrangel1126.github.io/outlookRCC/src/auth-redirect.html` — already registered, do not change |

---
Saved: April 28, 2026
