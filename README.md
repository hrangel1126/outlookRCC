# RCC Mail - Outlook Add-in

## Overview

Outlook Web Add-in to send emails from shared mailboxes.

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
- `src/compose.html/js` - Email compose form
- `src/settings.html/js` - Manage shared mailboxes
- `src/styles.css` - Styling

### Backend (Vercel)
- `vercel-backend/api/send-email.js` - Receives token, exchanges via OBO, sends email

### Manifest
- `manifest.xml` - Office add-in manifest

## Working Manifest Version

**IMPORTANT**: The manifest from commit `8d7229e` is the working version. Later commits that added:
- IdentityAPI requirement
- WebApplicationInfo section

...cause Exchange to reject sideloading with error: "Sideloading rejected by Exchange"

### Why This Happens
Adding WebApplicationInfo to enable Office SSO causes validation failures during sideload in some Exchange environments. The exact reason is unclear but may be related to:
- Azure app not being pre-approved by Exchange
- Resource URI format validation
- Missing admin consent in Azure AD

### To Enable SSO Later
When Office SSO is needed, you must:
1. Register the add-in in Azure AD properly
2. Get Exchange admin to allow sideloading
3. Add WebApplicationInfo back to manifest

## Authentication Flow

1. User clicks "Obtener Token" in the add-in
2. `Office.auth.getAccessToken()` retrieves token from user's existing Outlook session
3. Token is sent to Vercel backend
4. Vercel exchanges token for Graph API token (On-Behalf-Of flow)
5. Email is sent via Microsoft Graph API

**Note**: If Office SSO fails with error 13000, it means the identity API is not supported in this add-in context.

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

1. Use manifest from commit 8d7229e (or current working version)
2. Push code to GitHub (auto-deploys frontend)
3. Vercel auto-deploys backend
4. Sideload manifest.xml in Outlook

## Key Files

| File | Purpose |
|------|---------|
| manifest.xml | Office add-in manifest (use working version) |
| src/compose.js | Email compose with Office SSO |
| vercel-backend/api/send-email.js | Backend API with OBO token exchange |

## Known Issues

- Error 13000 = identity API not supported for this add-in
- WebApplicationInfo causes sideload rejection - DO NOT add to manifest
- Icons must point to GitHub Pages URLs (not Vercel)

## Session Lessons Learned

1. **Manifest Changes Break Sideloading**: Adding IdentityAPI and WebApplicationInfo to manifest causes "Sideloading rejected by Exchange" error
2. **Use Working Manifest Version**: Always use manifest from commit 8d7229e as base
3. **Icons Must Be on GitHub Pages**: Icon URLs in manifest must point to hrangel1126.github.io, not Vercel
4. **Office SSO Has Limitations**: Even with correct manifest, error 13000 can occur in certain Exchange environments