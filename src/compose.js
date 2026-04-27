// compose.js — Compose and send email via Microsoft Graph API
//
// Auth flow (MSAL.js popup):
//   1. First time → loginPopup() opens Microsoft login → user logs in once
//   2. Token saved to localStorage by MSAL → stays logged in across sessions
//   3. On send → acquireTokenSilent() (no popup) → POST to Graph API → email sent
//
// Sending logic:
//   - Selected "De" == personal account → POST /me/sendMail
//   - Selected "De" == shared mailbox   → POST /users/{mailbox}/sendMail
//     (requires user to have "Send As" rights on that mailbox in Exchange)

var msalInstance = null;
var loggedInAccount = null;
var userEmail = "";  // personal email of the logged-in user

Office.onReady(function () {
    msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
    loadMailboxes();
    checkAuthState();
});

// ── Auth ──────────────────────────────────────────────────────────────────────

// Check if user already has a cached token (from a previous session)
function checkAuthState() {
    var accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        setLoggedIn(accounts[0]);
    } else {
        showLoggedOut();
    }
}

// Show the login popup — user only needs to do this once
async function login() {
    try {
        showStatus("Abriendo ventana de Microsoft...", "info");
        var result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
        setLoggedIn(result.account);
        showStatus("Sesión iniciada correctamente.", "success");
    } catch (err) {
        // User closed the popup or there was an error
        if (err.errorCode !== "user_cancelled") {
            showStatus("Error al iniciar sesión: " + err.message, "error");
        } else {
            hideStatus();
        }
    }
}

function setLoggedIn(account) {
    loggedInAccount = account;
    userEmail = account.username.toLowerCase();
    document.getElementById("loggedInEmail").textContent = account.username;
    document.getElementById("loggedInBar").style.display = "block";
    document.getElementById("loggedOutBar").style.display = "none";
    document.getElementById("btnSend").disabled = false;
}

function showLoggedOut() {
    document.getElementById("loggedInBar").style.display = "none";
    document.getElementById("loggedOutBar").style.display = "block";
    document.getElementById("btnSend").disabled = true;
}

async function logout() {
    await msalInstance.logoutPopup({ account: loggedInAccount });
    loggedInAccount = null;
    userEmail = "";
    showLoggedOut();
    showStatus("Sesión cerrada.", "info");
}

// Get a valid Graph API token — silent from cache, popup only if expired/missing
async function getToken() {
    if (!loggedInAccount) throw new Error("No hay sesión activa. Inicie sesión primero.");

    try {
        var result = await msalInstance.acquireTokenSilent({
            scopes:  GRAPH_SCOPES,
            account: loggedInAccount
        });
        return result.accessToken;
    } catch (err) {
        // Token expired or needs new consent — show popup
        var result = await msalInstance.acquireTokenPopup({
            scopes:  GRAPH_SCOPES,
            account: loggedInAccount
        });
        return result.accessToken;
    }
}

// ── Send email ────────────────────────────────────────────────────────────────

async function sendEmail() {
    var toRaw = document.getElementById("toField").value.trim();
    if (!toRaw) {
        showStatus("El campo Para es requerido.", "error");
        return;
    }

    var fromMailbox = document.getElementById("fromSelect").value;
    var ccRaw       = document.getElementById("ccField").value.trim();
    var subject     = document.getElementById("subjectField").value.trim();
    var bodyText    = document.getElementById("bodyField").value;

    var toList = splitAddresses(toRaw);
    var ccList = ccRaw ? splitAddresses(ccRaw) : [];

    try {
        showStatus("Obteniendo token...", "info");
        var token = await getToken();

        // Personal account → /me/sendMail
        // Shared mailbox   → /users/{mailbox}/sendMail  (needs Send As rights)
        var isSharedMailbox = fromMailbox && fromMailbox.toLowerCase() !== userEmail;
        var endpoint = isSharedMailbox
            ? "https://graph.microsoft.com/v1.0/users/" + fromMailbox + "/sendMail"
            : "https://graph.microsoft.com/v1.0/me/sendMail";

        showStatus("Enviando...", "info");

        var response = await fetch(endpoint, {
            method:  "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type":  "application/json"
            },
            body: JSON.stringify(buildGraphPayload(toList, ccList, subject, bodyText))
        });

        if (response.status === 202) {
            showStatus("✓ Correo enviado desde " + (fromMailbox || userEmail), "success");
            clearForm();
        } else {
            // Graph returns error details in the body
            var err = await response.json().catch(function () { return {}; });
            var msg = (err.error && err.error.message) ? err.error.message : "HTTP " + response.status;
            showStatus("Error al enviar: " + msg, "error");
        }

    } catch (err) {
        showStatus("Error: " + err.message, "error");
    }
}

// Build the Graph API sendMail request body
function buildGraphPayload(toList, ccList, subject, bodyText) {
    return {
        message: {
            subject: subject,
            body: {
                contentType: "Text",
                content:     bodyText
            },
            toRecipients: toList.map(function (a) {
                return { emailAddress: { address: a } };
            }),
            ccRecipients: ccList.map(function (a) {
                return { emailAddress: { address: a } };
            })
        },
        saveToSentItems: true   // shows in the Sent folder of the sending mailbox
    };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

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
        opt.value = mb;
        opt.textContent = mb;
        if (mb === defaultEmail) opt.selected = true;
        select.appendChild(opt);
    });
}

function getMailboxes() {
    var json = localStorage.getItem("rcc_mailboxes");
    return json ? JSON.parse(json) : [];
}

// Split "a@x.com; b@x.com, c@x.com" → ["a@x.com", "b@x.com", "c@x.com"]
function splitAddresses(raw) {
    return raw.split(/[,;]/)
              .map(function (s) { return s.trim(); })
              .filter(Boolean);
}

function clearForm() {
    document.getElementById("toField").value      = "";
    document.getElementById("ccField").value      = "";
    document.getElementById("subjectField").value = "";
    document.getElementById("bodyField").value    = "";
}

function showStatus(msg, type) {
    var el = document.getElementById("statusMsg");
    el.textContent  = msg;
    el.className    = "status status-" + type;
    el.style.display = "block";
    if (type === "success") {
        setTimeout(function () { el.style.display = "none"; }, 5000);
    }
}

function hideStatus() {
    document.getElementById("statusMsg").style.display = "none";
}

function goBack() {
    window.location.href = "taskpane.html";
}
