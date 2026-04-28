// compose.js — Send email using Office SSO token

var VERCEL_API = "https://oulookrcc.vercel.app";
var API_KEY    = "rcc-api-key-2026";

var officeToken = null;

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        loadMailboxes();
    }
});

async function getToken() {
    if (!Office.auth || !Office.auth.getAccessToken) {
        showStatus("SSO no disponible en este contexto.", "error");
        return;
    }
    
    showStatus("Obteniendo token de Outlook...", "info");
    
    try {
        officeToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
        document.getElementById("tokenInput").value = officeToken;
        
        var payload = decodeJwt(officeToken);
        if (payload) {
            var expTime = payload.exp ? new Date(payload.exp * 1000).toLocaleTimeString() : "?";
            showStatus("✓ Token SSO | Expira: " + expTime, "success");
        } else {
            showStatus("✓ Token SSO obtenido", "success");
        }
    } catch (err) {
        showStatus("Error SSO (" + err.code + "): " + err.message, "error");
    }
}

async function sendEmail() {
    var toRaw = document.getElementById("toField").value.trim();
    if (!toRaw) {
        showStatus("El campo Para es requerido.", "error");
        return;
    }

    if (!officeToken) {
        var inputToken = document.getElementById("tokenInput").value.trim();
        if (inputToken) {
            officeToken = inputToken;
        } else {
            showStatus("Necesitas un token. Haz clic en Obtener Token.", "error");
            return;
        }
    }

    var fromMailbox = document.getElementById("fromSelect").value;
    var ccRaw       = document.getElementById("ccField").value.trim();
    var subject     = document.getElementById("subjectField").value.trim();
    var bodyText    = document.getElementById("bodyField").value;

    showStatus("Enviando...", "info");

    try {
        var response = await fetch(VERCEL_API + "/api/send-email", {
            method: "POST",
            headers: {
                "Content-Type":   "application/json",
                "x-api-key":       API_KEY,
                "x-office-token":  officeToken
            },
            body: JSON.stringify({
                from:    fromMailbox,
                to:      splitAddresses(toRaw),
                cc:      ccRaw ? splitAddresses(ccRaw) : [],
                subject: subject,
                body:    bodyText
            })
        });

        var result = await response.json();

        if (response.ok && result.success) {
            showStatus("✓ " + result.message, "success");
            clearForm();
            officeToken = null;
        } else {
            showStatus("Error: " + (result.detail || result.error), "error");
        }

    } catch (err) {
        showStatus("Error: " + err.message, "error");
    }
}

function loadMailboxes() {
    var mailboxes    = getMailboxes();
    var defaultEmail = localStorage.getItem("rcc_default_mailbox") || "";
    var select       = document.getElementById("fromSelect");

    select.innerHTML = "";

    if (mailboxes.length === 0) {
        var opt = document.createElement("option");
        opt.value = "";
        opt.textContent = "-- Configure un buzón en Configuracion --";
        select.appendChild(opt);
        return;
    }

    mailboxes.forEach(function (mb) {
        var opt = document.createElement("option");
        opt.value       = mb;
        opt.textContent = mb;
        if (mb === defaultEmail) opt.selected = true;
        select.appendChild(opt);
    });
}

function getMailboxes() {
    var json = localStorage.getItem("rcc_mailboxes");
    return json ? JSON.parse(json) : [];
}

function splitAddresses(raw) {
    return raw.split(/[,;]/).map(function(s){ return s.trim(); }).filter(Boolean);
}

function clearForm() {
    document.getElementById("toField").value      = "";
    document.getElementById("ccField").value      = "";
    document.getElementById("subjectField").value = "";
    document.getElementById("bodyField").value    = "";
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
    var el = document.getElementById("statusMsg");
    el.textContent   = msg;
    el.className     = "status status-" + type;
    el.style.display = "block";
    if (type === "success") setTimeout(function(){ el.style.display = "none"; }, 5000);
}

function goBack() { 
    if (Office.context && Office.context.ui) {
        Office.context.ui.closeContainer();
    } else {
        window.location.href = "taskpane.html";
    }
}