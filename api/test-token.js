// api/test-token.js — Receives an Office SSO token and returns its decoded claims
// Used for testing: confirms Office.auth.getAccessToken() works end-to-end
//
// POST /api/test-token
// Headers: x-api-key, x-office-token

export default async function handler(req, res) {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-api-key, x-office-token");

    if (req.method === "OPTIONS") return res.status(200).end();
    if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

    if (req.headers["x-api-key"] !== process.env.API_KEY) {
        return res.status(401).json({ error: "Unauthorized" });
    }

    const officeToken = req.headers["x-office-token"];
    if (!officeToken) return res.status(400).json({ error: "Missing x-office-token header" });

    try {
        // Decode JWT payload (no verification — this is just for debugging)
        const parts = officeToken.split(".");
        if (parts.length !== 3) return res.status(400).json({ error: "Invalid JWT format" });

        const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));

        return res.status(200).json({
            valid:    true,
            email:    payload.preferred_username || payload.upn || payload.unique_name || null,
            upn:      payload.upn || null,
            name:     payload.name || null,
            aud:      payload.aud || null,
            iss:      payload.iss || null,
            exp:      payload.exp || null,
            iat:      payload.iat || null,
            tid:      payload.tid || null,
            tokenPreview: officeToken.substring(0, 40) + "..."
        });

    } catch (err) {
        return res.status(500).json({ error: "Failed to decode token: " + err.message });
    }
}
