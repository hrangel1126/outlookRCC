// api/auth/login.js — Redirect to Microsoft login (used once to get refresh token)
// Open https://oulookrcc.vercel.app/api/auth/login in a browser to authorize

export default function handler(req, res) {
    const params = new URLSearchParams({
        client_id:     process.env.CLIENT_ID,
        response_type: "code",
        redirect_uri:  process.env.VERCEL_URL + "/api/auth/callback",
        scope:         "offline_access Mail.Send Mail.Send.Shared User.Read",
        response_mode: "query",
        prompt:        "select_account"
    });

    const authUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/authorize?${params}`;
    res.redirect(authUrl);
}
