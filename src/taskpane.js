// taskpane.js — Home panel
// Uses MSAL popup to get a Graph-scoped access token.
// Token stored in localStorage so send.html can use it directly.

var VERCEL_SEND_URL = "https://hrangel1126.github.io/outlookRCC/src/send.html";

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        showDefaultMailbox();
        checkStoredToken();
    }
});

function showDefaultMailbox() {
    var defaultMailbox = localStorage.getItem("rcc_default_mailbox") || "";
    var mailboxes = JSON.parse(localStorage.getItem("rcc_mailboxes") || "[]");

    if (defaultMailbox) {
        document.getElementById("defaultBadge").style.display = "block";
        document.getElementById("defaultName").textContent = defaultMailbox;
    }

    if (mailboxes.length === 0) {
        document.getElementById("noMailboxWarning").style.display = "block";
    }
}

// On load: show status of any stored token
function checkStoredToken() {
    var token = localStorage.getItem("rcc_office_token");
    if (!token) return;

    var payload = decodeJwt(token);
    if (!payload) return;

    var now = Math.floor(Date.now() / 1000);
    if (payload.exp && payload.exp > now) {
        var expTime = new Date(payload.exp * 1000).toLocaleTimeString();
        var user = payload.preferred_username || payload.upn || payload.unique_name || "";
        showStatus("✓ Token activo | " + user + " | Expira: " + expTime, "success");
    } else {
        showStatus("⚠ Token expirado — haz clic en Obtener Token para renovar.", "info");
    }
}

async function obtenerToken() {
    showStatus("Iniciando sesión...", "info");

    try {
        var pca      = new msal.PublicClientApplication(MSAL_CONFIG);
        var request  = { scopes: GRAPH_SCOPES };
        var accounts = pca.getAllAccounts();
        var result;

        if (accounts.length > 0) {
            // Silent refresh first — no popup if token is still cached
            try {
                result = await pca.acquireTokenSilent({ ...request, account: accounts[0] });
            } catch (silentErr) {
                // Silent failed (expired, no refresh token) → show popup
                result = await pca.acquireTokenPopup(request);
            }
        } else {
            // No cached account → show login popup
            result = await pca.acquireTokenPopup(request);
        }

        // Store raw access token for send.html to use directly
        localStorage.setItem("rcc_office_token", result.accessToken);

        var expTime = new Date(result.expiresOn).toLocaleTimeString();
        showStatus(
            "✓ Token obtenido | " + result.account.username + " | Expira: " + expTime +
            "\n\nAbre " + VERCEL_SEND_URL + " para enviar correos.",
            "success"
        );

    } catch (err) {
        showStatus("Error al obtener token: " + err.message, "error");
    }
}

function verificarEstado() {
    var token = localStorage.getItem("rcc_office_token");

    if (!token) {
        showStatus("Sin token almacenado. Haz clic en Obtener Token primero.", "info");
        return;
    }

    var payload = decodeJwt(token);
    if (!payload) {
        showStatus("Token almacenado inválido. Haz clic en Obtener Token.", "error");
        return;
    }

    var now     = Math.floor(Date.now() / 1000);
    var expTime = payload.exp ? new Date(payload.exp * 1000).toLocaleTimeString() : "?";
    var user    = payload.preferred_username || payload.upn || payload.unique_name || "?";

    if (payload.exp && payload.exp > now) {
        var minsLeft = Math.round((payload.exp - now) / 60);
        showStatus("✓ Token válido | " + user + " | Expira: " + expTime + " (" + minsLeft + " min restantes)", "success");
    } else {
        showStatus("⚠ Token expirado (expiró a las " + expTime + "). Haz clic en Obtener Token.", "error");
    }
}

function openSettings() {
    window.location.href = "settings.html";
}

function decodeJwt(token) {
    try {
        var parts = token.split(".");
        if (parts.length !== 3) return null;
        var json = atob(parts[1].replace(/-/g, "+").replace(/_/g, "/"));
        return JSON.parse(json);
    } catch (e) {
        return null;
    }
}

function showStatus(msg, type) {
    var el = document.getElementById("tokenStatus");
    el.textContent      = msg;
    el.className        = "status status-" + type;
    el.style.display    = "block";
    el.style.whiteSpace = "pre-line";
}
