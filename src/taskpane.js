// taskpane.js — Home panel
// Shows login status (token already in localStorage from compose page)
// and navigates to compose / settings.

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        showDefaultMailbox();
        showAuthStatus();
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

// Check MSAL localStorage cache to show who is logged in (read-only, no network call)
function showAuthStatus() {
    var pca = new msal.PublicClientApplication(MSAL_CONFIG);
    var accounts = pca.getAllAccounts();
    var el = document.getElementById("authStatus");

    if (accounts.length > 0) {
        el.textContent = "✓ " + accounts[0].username;
        el.className = "status status-success";
        el.style.display = "block";
    } else {
        el.textContent = "⚠ No hay sesión. Abre Enviar Correo para iniciar sesión.";
        el.className = "status status-info";
        el.style.display = "block";
    }
}

function openCompose() {
    window.location.href = "compose.html";
}

function openSettings() {
    window.location.href = "settings.html";
}
