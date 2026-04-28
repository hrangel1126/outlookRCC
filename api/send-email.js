// api/send-email.js — Receives email data + MSAL Graph token, sends via Graph API
//
// POST /api/send-email
// Headers:
//   x-api-key: YOUR_API_KEY     (protects this endpoint)
//   x-office-token: <JWT>       (MSAL access token scoped for Mail.Send / Mail.Send.Shared)
// Body: { from, to[], cc[], subject, body, isHtml? }

export default async function handler(req, res) {
    // CORS — allow requests from GitHub Pages (add-in host) and Vercel (send.html)
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-api-key, x-office-token");

    if (req.method === "OPTIONS") return res.status(200).end();
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    if (req.headers["x-api-key"] !== process.env.API_KEY) {
        return res.status(401).json({ error: "Unauthorized" });
    }

    const graphToken = req.headers["x-office-token"];
    if (!graphToken) return res.status(400).json({ error: "Missing token" });

    const { from, to, cc = [], subject, body, isHtml = false } = req.body;
    if (!to || to.length === 0) return res.status(400).json({ error: "Campo 'to' requerido" });

    try {
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
                    toRecipients: to.map(a => ({ emailAddress: { address: a } })),
                    ccRecipients: cc.map(a => ({ emailAddress: { address: a } }))
                },
                saveToSentItems: true
            })
        });

        if (graphRes.status === 202) {
            return res.status(200).json({ success: true, message: "Correo enviado desde " + (from || userEmail) });
        }

        const errBody = await graphRes.json().catch(() => ({}));
        return res.status(graphRes.status).json({ error: "Graph error", detail: errBody?.error?.message });

    } catch (err) {
        return res.status(500).json({ error: err.message });
    }
}

async function getUserEmail(graphToken) {
    const res  = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName", {
        headers: { "Authorization": "Bearer " + graphToken }
    });
    const data = await res.json();
    return (data.mail || data.userPrincipalName || "").toLowerCase();
}
