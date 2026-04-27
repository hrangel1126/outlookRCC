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
| Local server | Node.js HTTPS (server.js) | Not needed in VSTO |

No frameworks — plain HTML/JS. Easy to read and maintain.

---

## What the Add-in Does

1. **Home panel** (`taskpane.html`) — shows current default mailbox, two buttons
2. **Enviar Correo** (`compose.html`) — fill De/Para/CC/Asunto/Mensaje → opens Outlook compose window pre-filled
3. **Configuracion** (`settings.html`) — add/remove shared mailboxes, set default

Buttons appear on the **Home tab** of Outlook (TabDefault group "RCC Mail").
Also accessible from the **Apps sidebar** (left nav in New Outlook).

---

## Known Limitation

`Office.context.mailbox.displayNewMessageForm()` does NOT support setting the `From` field
programmatically — this is a current gap in the New Outlook API. The user sees a note
in the compose form and can change the From field manually in the Outlook compose window.

If Microsoft adds `from` support to displayNewMessageForm in the future, add it to
`compose.js` in the options object.

---

## File Structure

```
outlooknew/
├── manifest.xml          Office Add-in manifest — ribbon wiring
├── package.json          npm scripts: start, setup, install-certs
├── server.js             Local HTTPS server (port 3000) for testing
├── setup.js              Generates placeholder PNG icons in assets/
├── SESSION.md            This file
├── INSTALL.md            Step-by-step install instructions
├── src/
│   ├── styles.css        Shared CSS for all pages
│   ├── taskpane.html/js  Home panel (two buttons + default mailbox badge)
│   ├── compose.html/js   Compose email form
│   ├── settings.html/js  Manage mailboxes form
│   └── commands.html     Required manifest placeholder (no logic)
└── assets/
    ├── icon-16.png       Placeholder icons (blue squares)
    ├── icon-32.png       Replace with real branded icons before deploying
    ├── icon-64.png
    ├── icon-80.png
    └── icon-128.png
```

---

## To Continue Later

### Add Graph API (send directly without Outlook window)
Currently Enviar Correo opens the Outlook compose window (displayNewMessageForm).
To send directly (and set the From field), integrate MSAL.js + Microsoft Graph API:
- Register app in Azure → add Mail.Send.Shared permission
- Use `https://graph.microsoft.com/v1.0/users/{mailbox}/sendMail` with POST
- See the parent project at `C:\HR\RCCAPP` (RCCApp.vbproj) for Graph API / MSAL patterns

### Deploy to GitHub Pages (no local server needed)
1. Push the `outlooknew/` folder to a GitHub repo
2. Enable GitHub Pages on the repo
3. Replace all `https://localhost:3000` in manifest.xml with `https://yourusername.github.io/repo-name`
4. Sideload the updated manifest.xml

### Deploy to SharePoint / M365 Admin Center
For organization-wide deployment with zero user install steps:
1. Go to M365 Admin Center → Settings → Integrated apps → Upload custom apps
2. Upload manifest.xml
3. Add available to specific users or all users
4. Users see the add-in automatically in Outlook — no sideloading needed

---

## Lessons from the VSTO Project

- Use HKEY_CURRENT_USER — no admin rights required
- Duplicate ribbon files break the project (don't replicate that pattern here)
- .NET 4.7.2 / AnyCPU constraints do NOT apply to this web add-in
- Settings in localStorage do NOT roam across devices — if cross-device sync is needed,
  switch to `Office.context.roamingSettings` (Exchange-synced, max 32KB)

---
Saved: April 27, 2026
