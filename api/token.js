// api/token.js — Internal helper: get a fresh access token using the stored refresh token

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
    if (data.error) throw new Error("Token refresh failed: " + data.error_description);
    return data.access_token;
}
