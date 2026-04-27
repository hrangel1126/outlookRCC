// compose.js — Compose email form logic
//
// Flow: user fills the form → clicks Enviar →
//   Office.context.mailbox.displayNewMessageForm() opens Outlook's compose
//   window pre-filled with To, CC, Subject, Body.
//
// Limitation: the "From" / sender account cannot be set programmatically
// in New Outlook via Office.js. The user may need to change it manually
// in the compose window. Tracked for future fix if Microsoft adds the API.

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadMailboxes();
  }
});

// Populate the "De" dropdown with mailboxes saved in settings
function loadMailboxes() {
  var mailboxes = getMailboxes();
  var defaultMailbox = localStorage.getItem("rcc_default_mailbox") || "";
  var select = document.getElementById("fromSelect");

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
    if (mb === defaultMailbox) opt.selected = true;
    select.appendChild(opt);
  });
}

// Open Outlook compose window with fields pre-filled
function sendEmail() {
  var toRaw = document.getElementById("toField").value.trim();

  if (!toRaw) {
    showStatus("El campo Para es requerido.", "error");
    return;
  }

  var ccRaw     = document.getElementById("ccField").value.trim();
  var subject   = document.getElementById("subjectField").value.trim();
  var bodyText  = document.getElementById("bodyField").value;

  // Support multiple recipients separated by comma or semicolon
  var toList = splitAddresses(toRaw);
  var ccList = ccRaw ? splitAddresses(ccRaw) : [];

  // Convert plain-text line breaks to HTML so they display correctly
  var htmlBody = bodyText.replace(/&/g, "&amp;")
                         .replace(/</g, "&lt;")
                         .replace(/>/g, "&gt;")
                         .replace(/\n/g, "<br/>");

  try {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients:  toList,
      ccRecipients:  ccList,
      subject:       subject,
      htmlBody:      htmlBody
    });

    showStatus("Ventana de redacción abierta en Outlook.", "success");

    // Clear form after a short delay so the user can see the success message
    setTimeout(clearForm, 1500);

  } catch (err) {
    showStatus("Error al abrir la ventana de correo: " + err.message, "error");
  }
}

// Split "a@x.com; b@x.com, c@x.com" into ["a@x.com","b@x.com","c@x.com"]
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

function getMailboxes() {
  var json = localStorage.getItem("rcc_mailboxes");
  return json ? JSON.parse(json) : [];
}

function showStatus(msg, type) {
  var el = document.getElementById("statusMsg");
  el.textContent = msg;
  el.className = "status status-" + type;
  el.style.display = "block";
  setTimeout(function () { el.style.display = "none"; }, 4000);
}

function goBack() {
  window.location.href = "taskpane.html";
}
