// api/auth/login.js
// Step 1 of OAuth flow — redirect the user to Microsoft login.
// Visit this URL once to authorize the app and store the refresh token.
//
// Usage: open https://your-vercel.app/api/auth/login in a browser

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
