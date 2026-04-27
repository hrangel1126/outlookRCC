// api/token.js — Internal helper: get a fresh access token using the stored refresh token.
// The refresh token is stored in Vercel env vars (REFRESH_TOKEN).
// On each use, Microsoft issues a new refresh token — we update the env var automatically.
// If the refresh token expires (90 days of no use), re-run /api/auth/login.

export async function getFreshAccessToken() {
    const res = await fetch(
        `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
        {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({
                client_id:     process.env.CLIENT_ID,
                client_secret: process.env.CLIENT_SECRET,
                refresh_token: process.env.REFRESH_TOKEN,
                grant_type:    "refresh_token",
                scope:         "offline_access Mail.Send Mail.Send.Shared User.Read"
            })
        }
    );

    const data = await res.json();

    if (data.error) {
        throw new Error("Token refresh failed: " + data.error_description);
    }

    // Note: Microsoft returns a new refresh token on each call.
    // For long-term use, update REFRESH_TOKEN in Vercel env vars periodically,
    // or integrate Vercel KV to store it automatically.
    return data.access_token;
}
