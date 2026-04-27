// api/send-email.js — Main endpoint for sending email via Microsoft Graph API
//
// POST /api/send-email
// Headers: { "x-api-key": "your API_KEY env var" }
// Body (JSON):
// {
//   "from":    "buzon@empresa.com",   // shared mailbox or personal email
//   "to":      ["a@x.com", "b@x.com"],
//   "cc":      ["c@x.com"],           // optional
//   "subject": "Asunto",
//   "body":    "Cuerpo del mensaje",
//   "isHtml":  false                  // true for HTML body, false for plain text
// }
//
// Called from: the Outlook add-in, or any external script for bulk sending

import { getFreshAccessToken } from "./token.js";

export default async function handler(req, res) {
    // Only allow POST
    if (req.method !== "POST") {
        return res.status(405).json({ error: "Method not allowed" });
    }

    // Validate API key — protects the endpoint from unauthorized use
    const apiKey = req.headers["x-api-key"];
    if (!apiKey || apiKey !== process.env.API_KEY) {
        return res.status(401).json({ error: "Unauthorized" });
    }

    const { from, to, cc, subject, body, isHtml } = req.body;

    // Basic validation
    if (!to || to.length === 0) {
        return res.status(400).json({ error: "El campo 'to' es requerido." });
    }

    try {
        const accessToken = await getFreshAccessToken();

        // Determine endpoint:
        // If "from" is provided and different from the logged-in user → shared mailbox
        // The logged-in user must have "Send As" rights on the shared mailbox in Exchange
        const userEmail = await getLoggedInEmail(accessToken);
        const isShared  = from && from.toLowerCase() !== userEmail.toLowerCase();
        const endpoint  = isShared
            ? `https://graph.microsoft.com/v1.0/users/${from}/sendMail`
            : "https://graph.microsoft.com/v1.0/me/sendMail";

        const payload = buildPayload(to, cc || [], subject, body, isHtml || false);

        const graphRes = await fetch(endpoint, {
            method:  "POST",
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Content-Type":  "application/json"
            },
            body: JSON.stringify(payload)
        });

        if (graphRes.status === 202) {
            return res.status(200).json({
                success: true,
                message: "Correo enviado desde " + (from || userEmail)
            });
        }

        // Graph returned an error — pass it through
        const errBody = await graphRes.json().catch(() => ({}));
        return res.status(graphRes.status).json({
            error:  "Graph API error",
            detail: errBody.error ? errBody.error.message : "HTTP " + graphRes.status
        });

    } catch (err) {
        return res.status(500).json({ error: err.message });
    }
}

// Get the email address of the authenticated user
async function getLoggedInEmail(accessToken) {
    const res  = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName", {
        headers: { "Authorization": "Bearer " + accessToken }
    });
    const data = await res.json();
    return (data.mail || data.userPrincipalName || "").toLowerCase();
}

function buildPayload(to, cc, subject, body, isHtml) {
    return {
        message: {
            subject: subject || "(sin asunto)",
            body: {
                contentType: isHtml ? "HTML" : "Text",
                content:     body || ""
            },
            toRecipients: to.map(function(a) {
                return { emailAddress: { address: a.trim() } };
            }),
            ccRecipients: cc.map(function(a) {
                return { emailAddress: { address: a.trim() } };
            })
        },
        saveToSentItems: true
    };
}
