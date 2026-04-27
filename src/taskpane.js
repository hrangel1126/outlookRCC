// taskpane.js — Home panel logic
// Shows the default mailbox name and navigation buttons.

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    showDefaultMailbox();
  }
});

// Read saved default mailbox from localStorage and display it on screen.
// localStorage persists between sessions inside the add-in webview.
function showDefaultMailbox() {
  var defaultMailbox = localStorage.getItem("rcc_default_mailbox") || "";
  var mailboxes = getMailboxes();

  if (defaultMailbox) {
    document.getElementById("defaultBadge").style.display = "block";
    document.getElementById("defaultName").textContent = defaultMailbox;
  }

  // Warn the user if no mailboxes have been configured at all
  if (mailboxes.length === 0) {
    document.getElementById("noMailboxWarning").style.display = "block";
  }
}

function getMailboxes() {
  var json = localStorage.getItem("rcc_mailboxes");
  return json ? JSON.parse(json) : [];
}

// Navigate to compose form inside the task pane
function openCompose() {
  window.location.href = "compose.html";
}

// Navigate to settings form inside the task pane
function openSettings() {
  window.location.href = "settings.html";
}
