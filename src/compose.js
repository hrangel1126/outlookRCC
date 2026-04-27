// compose.js — Compose and send email via Office SSO + Vercel backend
//
// Flow:
//   1. Office.auth.getAccessToken() — borrows the user's existing Outlook session
//      No popup, no login screen — user is already signed into Outlook
//   2. Token sent to Vercel backend (/api/send-email)
//   3. Vercel exchanges token for Graph API token (OBO flow) and sends the email
//
// Update VERCEL_API and API_KEY after deploying to Vercel

var VERCEL_API = "https://your-app.vercel.app"; // ← replace with your Vercel URL
var API_KEY    = "your-api-key-here";            // ← must match API_KEY in Vercel env vars

Office.onReady(function () {
    loadMailboxes();
});

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

    try {
        showStatus("Obteniendo sesión de Outlook...", "info");

        // Get token from the user's existing Outlook session — no popup
        var officeToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });

        showStatus("Enviando...", "info");

        var response = await fetch(VERCEL_API + "/api/send-email", {
            method: "POST",
            headers: {
                "Content-Type":  "application/json",
                "x-api-key":     API_KEY,
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
        } else {
            showStatus("Error: " + (result.detail || result.error), "error");
        }

    } catch (err) {
        // 13xxx errors are Office SSO errors — show a clear message
        if (err.code) {
            showStatus("Error de sesión Outlook (" + err.code + "): " + err.message, "error");
        } else {
            showStatus("Error: " + err.message, "error");
        }
    }
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

function showStatus(msg, type) {
    var el = document.getElementById("statusMsg");
    el.textContent   = msg;
    el.className     = "status status-" + type;
    el.style.display = "block";
    if (type === "success") setTimeout(function(){ el.style.display = "none"; }, 5000);
}

function goBack() { window.location.href = "taskpane.html"; }
