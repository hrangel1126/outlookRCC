# Expected Work — RCC Mail Add-in Flow

## Overview

The add-in is installed once in Outlook via `manifest.xml` (sideload).
After that, the "RCC Mail" buttons appear permanently in the ribbon.
Sending an email goes through 5 steps:

---

## Step 1 — Install the add-in (one time)

Sideload `manifest.xml` into New Outlook or Outlook Web.
The "RCC Mail" group appears in the Home ribbon with two buttons:
- **Enviar Correo** — compose and send
- **Configuracion** — manage shared mailboxes

---

## Step 2 — Get an Office SSO token

When the user clicks **Enviar Correo** and hits **Enviar**, the add-in calls:

```js
Office.auth.getAccessToken()
```

This borrows the session the user **already has open in Outlook** — no extra login popup.
Outlook returns a short-lived JWT token (valid ~1 hour).

---

## Step 3 — Send token + email data to Vercel

The add-in posts to the Vercel backend:

```
POST https://oulookrcc.vercel.app/api/send-email
Headers:
  x-office-token: <JWT from Outlook>
  x-api-key: rcc-api-key-2026
Body: { from, to[], cc[], subject, body }
```

The Office token is proof of identity.
The `x-api-key` protects the endpoint from unauthorized callers.

---

## Step 4 — Vercel exchanges the token (OBO flow)

Vercel talks to Microsoft using the **On-Behalf-Of (OBO)** flow:

> "Here is a token from this user — give me a Graph API token so I can send email on their behalf."

Microsoft validates it and returns a Graph API access token.
The `CLIENT_SECRET` lives only in Vercel environment variables — never in the add-in code.

---

## Step 5 — Email is sent via Microsoft Graph API

Vercel uses the Graph token to call:

```
POST https://graph.microsoft.com/v1.0/users/{shared-mailbox}/sendMail
Authorization: Bearer <graph token>
```

The email is sent from the selected shared mailbox.
The user never left Outlook.

---

## Key Points

- The Office token from Outlook is **not** used directly to send email — it is only proof of identity.
- Vercel is the one that converts it into a real Graph API token and does the actual sending.
- Shared mailbox sending requires the `Mail.Send.Shared` permission granted in Azure App Registration.
- If the user's mailbox matches the `from` address, it uses `/me/sendMail` instead.
