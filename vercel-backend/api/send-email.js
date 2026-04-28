// api/send-email.js — Receives Office SSO token, exchanges via OBO, sends email
//
// POST /api/send-email
// Headers:
//   x-api-key: YOUR_API_KEY   (protects this endpoint)
//   x-office-token: <JWT>     (from Office.auth.getAccessToken())
// Body: { from, to[], cc[], subject, body, isHtml? }

export default async function handler(req, res) {
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    // Validate API key
    if (req.headers["x-api-key"] !== process.env.API_KEY) {
        return res.status(401).json({ error: "Unauthorized" });
    }

    const officeToken = req.headers["x-office-token"];
    if (!officeToken) return res.status(400).json({ error: "Falta el token de Office SSO" });

    const { from, to, cc = [], subject, body, isHtml = false } = req.body;
    if (!to || to.length === 0) return res.status(400).json({ error: "Campo 'to' requerido" });

    try {
        // OBO: exchange Office SSO token for Graph API token
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

// OBO: exchange Office SSO token for Graph API token
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