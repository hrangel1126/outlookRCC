// compose.js — Compose and send email via Office SSO or MSAL + Vercel backend

var VERCEL_API = "https://oulookrcc.vercel.app";
var API_KEY    = "rcc-api-key-2026";

var officeToken = null;

var MSAL_CONFIG = {
    auth: {
        clientId:    "870de84d-3b21-449c-bf57-4cb3c76f9893",
        authority:   "https://login.microsoftonline.com/common",
        redirectUri: window.location.href
    },
    cache: {
        cacheLocation:          "sessionStorage",
        storeAuthStateInCookie: false
    }
};

var GRAPH_SCOPES = ["Mail.Send", "Mail.Send.Shared", "User.Read"];

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        loadMailboxes();
    }
});

async function pasteToken() {
    try {
        var text = await navigator.clipboard.readText();
        document.getElementById("tokenInput").value = text;
        officeToken = text;
        showStatus("Token pegado desde portapapeles", "success");
    } catch (err) {
        showStatus("Error al pegar: " + err.message, "error");
    }
}

async function loginMsal() {
    showStatus("Iniciando sesión...", "info");
    
    if (typeof msal === "undefined") {
        showStatus("Cargando MSAL...", "info");
        // Wait for MSAL to load
        await new Promise(function(resolve) {
            setTimeout(function() {
                if (typeof msal !== "undefined") resolve();
                else resolve();
            }, 2000);
        });
        
        if (typeof msal === "undefined") {
            showStatus("Error: MSAL no se cargó. Recarga la página.", "error");
            return;
        }
    }
    
    try {
        var pca = new msal.PublicClientApplication(MSAL_CONFIG);
        var accounts = pca.getAllAccounts();
        
        var result;
        if (accounts.length > 0) {
            try {
                result = await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: accounts[0] });
            } catch (e) {
                result = await pca.acquireTokenPopup({ scopes: GRAPH_SCOPES });
            }
        } else {
            result = await pca.acquireTokenPopup({ scopes: GRAPH_SCOPES });
        }
        
        officeToken = result.accessToken;
        document.getElementById("tokenInput").value = officeToken;
        
        var expTime = result.expiresOn ? result.expiresOn.toLocaleTimeString() : "?";
        showStatus("✓ Sesión iniciada | Expira: " + expTime, "success");
        
    } catch (err) {
        showStatus("Error login: " + err.message, "error");
    }
}

document.getElementById("tokenInput").addEventListener("input", function() {
    officeToken = this.value.trim();
});

async function getToken() {
    // First try Office SSO
    if (Office.auth && Office.auth.getAccessToken) {
        try {
            showStatus("Obteniendo token SSO...", "info");
            officeToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
            document.getElementById("tokenInput").value = officeToken;
            
            var payload = decodeJwt(officeToken);
            if (payload) {
                var expTime = payload.exp ? new Date(payload.exp * 1000).toLocaleTimeString() : "?";
                showStatus("✓ Token SSO | Expira: " + expTime, "success");
            } else {
                showStatus("✓ Token SSO obtenido", "success");
            }
            return;
        } catch (err) {
            if (err.code !== 13000) {
                showStatus("Error SSO (" + err.code + "): " + err.message, "error");
                return;
            }
            // Fall through to MSAL
        }
    }
    
    // Fallback: MSAL popup login
    showStatus("SSO no disponible. Intentando login...", "info");
    await loginMsal();
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
            showStatus("Necesitas un token. Usa 'Obtener Token' o pégalo.", "error");
            return;
        }
    }
});

async function getToken() {
    showStatus("Obteniendo token de Outlook...", "info");
    
    try {
        officeToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
        
        var payload = decodeJwt(officeToken);
        if (payload) {
            var expTime = payload.exp ? new Date(payload.exp * 1000).toLocaleTimeString() : "?";
            showStatus("✓ Token obtenido | Expira: " + expTime, "success");
        } else {
            showStatus("✓ Token obtenido", "success");
        }
    } catch (err) {
        if (err.code === 13000) {
            showStatus("SSO no disponible en este contexto.", "info");
        } else {
            showStatus("Error SSO (" + err.code + "): " + err.message, "error");
        }
    }
}

async function sendEmail() {
    var toRaw = document.getElementById("toField").value.trim();
    if (!toRaw) {
        showStatus("El campo Para es requerido.", "error");
        return;
    }

    if (!officeToken) {
        await getToken();
        if (!officeToken) {
            showStatus("No hay token. Intenta de nuevo.", "error");
            return;
        }
    }

    var fromMailbox = document.getElementById("fromSelect").value;
    var ccRaw       = document.getElementById("ccField").value.trim();
    var subject     = document.getElementById("subjectField").value.trim();
    var bodyText    = document.getElementById("bodyField").value;

    try {
        showStatus("Enviando...", "info");

        var response = await fetch(VERCEL_API + "/api/send-email", {
            method: "POST",
            headers: {
                "Content-Type":   "application/json",
                "x-api-key":      API_KEY,
                "x-office-token": officeToken
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
        if (err.code) {
            showStatus("Error Outlook (" + err.code + "): " + err.message, "error");
        } else {
            showStatus("Error: " + err.message, "error");
        }
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