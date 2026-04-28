// taskpane.js — Home panel
// Obtains the Office SSO token and shows status.
// Actual email sending happens at https://oulookrcc.vercel.app/send.html

var VERCEL_SEND_URL = "https://oulookrcc.vercel.app/src/send.html";

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

// On load: check if a stored token exists and whether it is still valid
function checkStoredToken() {
    var token = localStorage.getItem("rcc_office_token");
    if (!token) return;

    var payload = decodeJwt(token);
    if (!payload) return;

    var now = Math.floor(Date.now() / 1000);
    if (payload.exp && payload.exp > now) {
        var expTime = new Date(payload.exp * 1000).toLocaleTimeString();
        showStatus("✓ Token activo | " + (payload.preferred_username || payload.upn || "") + " | Expira: " + expTime, "success");
    } else {
        showStatus("⚠ Token expirado — haz clic en Obtener Token para renovar.", "info");
    }
}

async function obtenerToken() {
    try {
        showStatus("Obteniendo token de Outlook...", "info");

        var token = await Office.auth.getAccessToken({ allowSignInPrompt: true });
        localStorage.setItem("rcc_office_token", token);

        var payload = decodeJwt(token);
        var user    = payload ? (payload.preferred_username || payload.upn || "?") : "?";
        var expTime = payload && payload.exp ? new Date(payload.exp * 1000).toLocaleTimeString() : "?";

        showStatus(
            "✓ Token obtenido | " + user + " | Expira: " + expTime +
            "\n\nAbre " + VERCEL_SEND_URL + " para enviar correos.",
            "success"
        );
    } catch (err) {
        var msg = err.code ? "Error SSO (" + err.code + "): " + err.message : "Error: " + err.message;
        showStatus(msg, "error");
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
    var user    = payload.preferred_username || payload.upn || "?";

    if (payload.exp && payload.exp > now) {
        var minsLeft = Math.round((payload.exp - now) / 60);
        showStatus("✓ Token válido | " + user + " | Expira: " + expTime + " (" + minsLeft + " min restantes)", "success");
    } else {
        showStatus("⚠ Token expirado (expiró a las " + expTime + "). Haz clic en Obtener Token para renovar.", "error");
    }
}

function openSettings() {
    window.location.href = "settings.html";
}

// Decode a JWT payload without verifying the signature
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
    el.textContent   = msg;
    el.className     = "status status-" + type;
    el.style.display = "block";
    el.style.whiteSpace = "pre-line";
}
