# RCC Mail - Outlook Add-in

## Overview

Outlook Web Add-in to send emails from shared mailboxes using Office SSO authentication.

## Architecture

```
┌─────────────────┐     ┌──────────────┐     ┌─────────────────┐     ┌─────────────┐
│   Outlook       │     │  GitHub      │     │    Vercel       │     │   Microsoft │
│   Add-in        │────▶│  Pages       │────▶│   Backend       │────▶│   Graph API │
│  (Office.js)   │     │  (Frontend)  │     │   (API)         │     │             │
└─────────────────┘     └──────────────┘     └─────────────────┘     └─────────────┘
        │                                           │
        │ Office SSO                                │ OBO Exchange
        │ Token                                     │ (Azure App)
        ▼                                           ▼
┌─────────────────┐                        ┌─────────────────┐
│ User's Outlook  │                        │   Azure AD      │
│   Session       │                        │ (App Registry)  │
└─────────────────┘                        └─────────────────┘
```

## Components

### Frontend (GitHub Pages)
- `src/taskpane.html/js` - Home panel
- `src/compose.html/js` - Email compose form (uses Office SSO)
- `src/settings.html/js` - Manage shared mailboxes
- `src/styles.css` - Styling

### Backend (Vercel)
- `vercel-backend/api/send-email.js` - Receives token, exchanges via OBO, sends email

### Manifest
- `manifest.xml` - Office add-in manifest with SSO configuration

## Authentication Flow

1. User clicks "Obtener Token" in the add-in
2. `Office.auth.getAccessToken()` retrieves token from user's existing Outlook session
3. Token is sent to Vercel backend
4. Vercel exchanges token for Graph API token (On-Behalf-Of flow)
5. Email is sent via Microsoft Graph API

## Azure App Registration

Created in Microsoft Entra ID:

- **CLIENT_ID**: `3b387ca6-bd43-4396-b557-bbd5786405db`
- **TENANT_ID**: `98510037-46a2-4ae4-85d2-9130a24f7af1`
- **CLIENT_SECRET**: (stored in Vercel)

### API Permissions (Delegated)
- Mail.Send
- Mail.Send.Shared
- User.Read

## Vercel Environment Variables

```
CLIENT_ID=3b387ca6-bd43-4396-b557-bbd5786405db
CLIENT_SECRET=<secret>
TENANT_ID=98510037-46a2-4ae4-85d2-9130a24f7af1
API_KEY=rcc-api-key-2026
```

## URLs

- **Frontend**: https://hrangel1126.github.io/outlookRCC/
- **Backend**: https://oulookrcc.vercel.app

## Deployment

1. Push code to GitHub (auto-deploys frontend)
2. Vercel auto-deploys backend
3. Sideload manifest.xml in Outlook

## Key Files

| File | Purpose |
|------|---------|
| manifest.xml | Office add-in manifest with SSO config |
| src/compose.js | Email compose with Office SSO |
| vercel-backend/api/send-email.js | Backend API with OBO token exchange |

## Known Issues

- Office SSO requires IdentityAPI requirement in manifest
- WebApplicationInfo must match the Azure app CLIENT_ID
- Error 13000 = identity API not supported (personal account or SSO unavailable)