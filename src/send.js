// send.js — Standalone send page at https://oulookrcc.vercel.app/send.html
// Does NOT require Office.js — works in any browser tab.
// Token is obtained from the Outlook add-in (taskpane) and stored in localStorage.

var VERCEL_API = "https://oulookrcc.vercel.app";
var API_KEY    = "rcc-api-key-2026";

window.addEventListener("load", function () {
    // Auto-fill token from localStorage if available
    var stored = localStorage.getItem("rcc_office_token");
    if (stored) {
        document.getElementById("tokenField").value = stored;
        showTokenInfo(stored);
    }

    // Auto-fill default mailbox if set (same localStorage, same origin on Vercel)
    var defaultMailbox = localStorage.getItem("rcc_default_mailbox");
    if (defaultMailbox) {
        document.getElementById("fromField").value = defaultMailbox;
    }
});

// Update token info whenever the user edits the token field
document.addEventListener("DOMContentLoaded", function () {
    document.getElementById("tokenField").addEventListener("input", function () {
        showTokenInfo(this.value.trim());
    });
});

async function sendEmail() {
    var token   = document.getElementById("tokenField").value.trim();
    var fromVal = document.getElementById("fromField").value.trim();
    var toRaw   = document.getElementById("toField").value.trim();
    var ccRaw   = document.getElementById("ccField").value.trim();
    var subject = document.getElementById("subjectField").value.trim();
    var body    = document.getElementById("bodyField").value;

    if (!token) { showStatus("Pega el token de Outlook primero.", "error"); return; }
    if (!toRaw) { showStatus("El campo Para es requerido.", "error"); return; }

    showStatus("Enviando...", "info");

    try {
        var response = await fetch(VERCEL_API + "/api/send-email", {
            method: "POST",
            headers: {
                "Content-Type":   "application/json",
                "x-api-key":      API_KEY,
                "x-office-token": token
            },
            body: JSON.stringify({
                from:    fromVal || undefined,
                to:      splitAddresses(toRaw),
                cc:      ccRaw ? splitAddresses(ccRaw) : [],
                subject: subject,
                body:    body
            })
        });

        var result = await response.json();

        if (response.ok && result.success) {
            showStatus("✓ " + result.message, "success");
        } else {
            showStatus("Error: " + (result.detail || result.error), "error");
        }
    } catch (err) {
        showStatus("Error de red: " + err.message, "error");
    }
}

function copyToken() {
    var val = document.getElementById("tokenField").value;
    if (!val) return;
    navigator.clipboard.writeText(val).then(function () {
        showStatus("Token copiado al portapapeles.", "info");
    });
}

function clearForm() {
    document.getElementById("toField").value      = "";
    document.getElementById("ccField").value      = "";
    document.getElementById("subjectField").value = "";
    document.getElementById("bodyField").value    = "";
    document.getElementById("statusMsg").style.display = "none";
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function splitAddresses(raw) {
    return raw.split(/[,;]/).map(function (s) { return s.trim(); }).filter(Boolean);
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

function showTokenInfo(token) {
    var el = document.getElementById("tokenInfo");
    if (!token) { el.textContent = ""; return; }

    var payload = decodeJwt(token);
    if (!payload) { el.textContent = "Token inválido."; return; }

    var now     = Math.floor(Date.now() / 1000);
    var user    = payload.preferred_username || payload.upn || "?";
    var expTime = payload.exp ? new Date(payload.exp * 1000).toLocaleTimeString() : "?";
    var valid   = payload.exp && payload.exp > now;

    el.textContent = (valid ? "✓ Válido" : "⚠ Expirado") + " | " + user + " | Expira: " + expTime;
    el.style.color = valid ? "#107c10" : "#a4262c";
}

function showStatus(msg, type) {
    var el = document.getElementById("statusMsg");
    el.textContent   = msg;
    el.className     = "status status-" + type;
    el.style.display = "block";
    if (type === "success") setTimeout(function () { el.style.display = "none"; }, 5000);
}
