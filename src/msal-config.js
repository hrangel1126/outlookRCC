// msal-config.js — Shared MSAL configuration for the RCC Mail add-in
//
// Reuses the same Azure App Registration as RCCApp (C:\HR\RCCAPP\AuthConfig.vb).
// CLIENT_ID and TENANT_ID are identical — no new registration needed.
//
// The only addition needed in Azure Portal (one time):
//   Authentication → Add platform → Single-page application
//   Redirect URI: https://hrangel1126.github.io/outlookRCC/src/auth-redirect.html

const MSAL_CONFIG = {
    auth: {
        clientId:    "870de84d-3b21-449c-bf57-4cb3c76f9893",
        authority:   "https://login.microsoftonline.com/a11b8361-d78a-47e9-8795-3d03ba2109c7",
        redirectUri: "https://hrangel1126.github.io/outlookRCC/src/auth-redirect.html"
    },
    cache: {
        // localStorage keeps the token across sessions — user stays logged in
        cacheLocation:       "localStorage",
        storeAuthStateInCookie: false
    }
};

// Scopes mirror what RCCApp requests.
// Mail.Send.Shared allows sending AS a shared mailbox (not just on behalf).
const GRAPH_SCOPES = [
    "Mail.Send",
    "Mail.Send.Shared",
    "User.Read"
];
