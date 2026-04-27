// compose.js — Compose and send email
//
// Sends via Vercel backend (POST /api/send-email).
// The backend holds the Microsoft token — the add-in never handles auth directly.
// Only an API key is needed here to call the backend endpoint.
//
// To update the Vercel URL: change VERCEL_API below.

var VERCEL_API = "https://your-app.vercel.app"; // ← replace after Vercel deploy
var API_KEY    = "change-this-to-match-vercel-env"; // ← must match API_KEY in Vercel

Office.onReady(function () {
    loadMailboxes();
});

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
        showStatus("Enviando...", "info");

        var response = await fetch(VERCEL_API + "/api/send-email", {
            method:  "POST",
            headers: {
                "Content-Type": "application/json",
                "x-api-key":    API_KEY
            },
            body: JSON.stringify({
                from:    fromMailbox,
                to:      toList,
                cc:      ccList,
                subject: subject,
                body:    bodyText,
                isHtml:  false
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
        showStatus("Error de conexión: " + err.message, "error");
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
    el.textContent   = msg;
    el.className     = "status status-" + type;
    el.style.display = "block";
    if (type === "success") {
        setTimeout(function () { el.style.display = "none"; }, 5000);
    }
}

function goBack() {
    window.location.href = "taskpane.html";
}
