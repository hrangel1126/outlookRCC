// api/auth/callback.js — Microsoft redirects here after login; shows the REFRESH_TOKEN

export default async function handler(req, res) {
    const { code, error } = req.query;

    if (error) return res.status(400).send("Auth error: " + error);
    if (!code)  return res.status(400).send("No auth code received.");

    const tokenRes = await fetch(
        `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
        {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({
                client_id:     process.env.CLIENT_ID,
                client_secret: process.env.CLIENT_SECRET,
                code:          code,
                redirect_uri:  process.env.VERCEL_URL + "/api/auth/callback",
                grant_type:    "authorization_code",
                scope:         "offline_access Mail.Send Mail.Send.Shared User.Read"
            })
        }
    );

    const tokens = await tokenRes.json();
    if (tokens.error) {
        return res.status(400).json({ error: tokens.error, detail: tokens.error_description });
    }

    res.send(`
        <html><body style="font-family:sans-serif; padding:20px;">
        <h2>✅ Login exitoso</h2>
        <p>Copia este <strong>REFRESH_TOKEN</strong> en tus variables de entorno de Vercel:</p>
        <textarea style="width:100%;height:120px;font-size:11px;">${tokens.refresh_token}</textarea>
        <p>Una vez guardado en Vercel, ya puedes cerrar esta página.</p>
        </body></html>
    `);
}
