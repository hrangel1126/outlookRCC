// msal-config.js — Shared MSAL configuration for the RCC Mail add-in
//
// App Registration is owned by RCC (Hebert's Azure) — clients do NOT need
// to register anything in their own Azure.
//
// How it works for client users:
//   1. Client opens the add-in → clicks "Iniciar sesión"
//   2. Standard Microsoft login popup appears (Microsoft's own UI)
//   3. Client logs in with THEIR Microsoft/M365 account
//   4. Microsoft asks them to accept permissions once (Mail.Send, User.Read)
//   5. Token saved to localStorage — they stay logged in forever
//
// Authority is "common" (not tenant-specific) so ANY Microsoft org or
// personal account can log in, not just accounts in RCC's tenant.

const MSAL_CONFIG = {
    auth: {
        clientId:    "870de84d-3b21-449c-bf57-4cb3c76f9893",
        authority:   "https://login.microsoftonline.com/common",
        redirectUri: "https://oulookrcc.vercel.app/src/auth-redirect.html"
    },
    cache: {
        // localStorage keeps the token across sessions — user stays logged in
        cacheLocation:          "localStorage",
        storeAuthStateInCookie: false
    }
};

// Mail.Send       → send from personal account
// Mail.Send.Shared → send from/as a shared mailbox
// User.Read       → get logged-in user's email address
const GRAPH_SCOPES = [
    "Mail.Send",
    "Mail.Send.Shared",
    "User.Read"
];
