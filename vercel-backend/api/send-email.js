// api/send-email.js — Receives email data + Office SSO token, sends via Graph API
//
// POST /api/send-email
// Headers:
//   x-api-key: YOUR_API_KEY        (protects this endpoint)
//   x-office-token: <JWT>          (from Office.auth.getAccessToken() in the add-in)
// Body: { from, to[], cc[], subject, body, isHtml? }

export default async function handler(req, res) {
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    // Validate API key
    if (req.headers["x-api-key"] !== process.env.API_KEY) {
        return res.status(401).json({ error: "Unauthorized" });
    }

    const officeToken = req.headers["x-office-token"];
    if (!officeToken) return res.status(400).json({ error: "Missing Office SSO token" });

    const { from, to, cc = [], subject, body, isHtml = false } = req.body;
    if (!to || to.length === 0) return res.status(400).json({ error: "Campo 'to' requerido" });

    try {
        // Exchange the Office SSO token for a Microsoft Graph token via OBO flow
        const graphToken = await exchangeForGraphToken(officeToken);

        // Determine endpoint — personal vs shared mailbox
        const userEmail = await getUserEmail(graphToken);
        const isShared  = from && from.toLowerCase() !== userEmail.toLowerCase();
        const endpoint  = isShared
            ? `https://graph.microsoft.com/v1.0/users/${from}/sendMail`
            : "https://graph.microsoft.com/v1.0/me/sendMail";

        const graphRes = await fetch(endpoint, {
            method:  "POST",
            headers: { "Authorization": "Bearer " + graphToken, "Content-Type": "application/json" },
            body: JSON.stringify({
                message: {
                    subject: subject || "(sin asunto)",
                    body:    { contentType: isHtml ? "HTML" : "Text", content: body || "" },
                    toRecipients: to.map(a  => ({ emailAddress: { address: a } })),
                    ccRecipients: cc.map(a  => ({ emailAddress: { address: a } }))
                },
                saveToSentItems: true
            })
        });

        if (graphRes.status === 202) {
            return res.status(200).json({ success: true, message: "Correo enviado desde " + (from || userEmail) });
        }

        const err = await graphRes.json().catch(() => ({}));
        return res.status(graphRes.status).json({ error: "Graph error", detail: err?.error?.message });

    } catch (err) {
        return res.status(500).json({ error: err.message });
    }
}

// OBO: exchange the Office SSO token for a Graph API token
// Uses CLIENT_ID + CLIENT_SECRET from Vercel environment variables
async function exchangeForGraphToken(officeToken) {
    const res = await fetch(
        `https://login.microsoftonline.com/${process.env.TENANT_ID || "common"}/oauth2/v2.0/token`,
        {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({
                grant_type:            "urn:ietf:params:oauth:grant-type:jwt-bearer",
                client_id:             process.env.CLIENT_ID,
                client_secret:         process.env.CLIENT_SECRET,
                assertion:             officeToken,
                requested_token_use:   "on_behalf_of",
                scope:                 "Mail.Send Mail.Send.Shared User.Read"
            })
        }
    );
    const data = await res.json();
    if (data.error) throw new Error(data.error_description || data.error);
    return data.access_token;
}

async function getUserEmail(graphToken) {
    const res  = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName", {
        headers: { "Authorization": "Bearer " + graphToken }
    });
    const data = await res.json();
    return (data.mail || data.userPrincipalName || "").toLowerCase();
}
