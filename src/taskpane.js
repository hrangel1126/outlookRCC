// taskpane.js — Home panel for RCC Mail add-in
// Uses Office SSO to get token (no MSAL needed)

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        showDefaultMailbox();
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

function openSettings() {
    window.location.href = "settings.html";
}