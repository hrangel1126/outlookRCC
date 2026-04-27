// settings.js — Manage shared mailboxes
//
// Storage (localStorage — survives across sessions in the add-in webview):
//   rcc_mailboxes       → JSON array of email strings
//   rcc_default_mailbox → string, the currently selected default

var selectedMailbox = null; // tracks which list item the user clicked

Office.onReady(function () {
  renderList();
});

// --- Storage helpers ---

function getMailboxes() {
  var json = localStorage.getItem("rcc_mailboxes");
  return json ? JSON.parse(json) : [];
}

function saveMailboxes(list) {
  localStorage.setItem("rcc_mailboxes", JSON.stringify(list));
}

function getDefault() {
  return localStorage.getItem("rcc_default_mailbox") || "";
}

function saveDefault(email) {
  localStorage.setItem("rcc_default_mailbox", email);
}

// --- List rendering ---

function renderList() {
  var mailboxes    = getMailboxes();
  var defaultEmail = getDefault();
  var list         = document.getElementById("mailboxList");

  list.innerHTML = "";
  selectedMailbox = null;

  if (mailboxes.length === 0) {
    list.innerHTML = '<li class="empty-msg">Sin buzones configurados</li>';
    return;
  }

  mailboxes.forEach(function (mb) {
    var li = document.createElement("li");
    li.textContent = mb;
    li.dataset.email = mb;

    // Mark the default mailbox with a star
    if (mb === defaultEmail) li.classList.add("is-default");

    li.addEventListener("click", function () {
      // Deselect previous, select this one
      list.querySelectorAll("li").forEach(function (el) {
        el.classList.remove("selected");
      });
      li.classList.add("selected");
      selectedMailbox = mb;
    });

    list.appendChild(li);
  });
}

// --- Actions ---

// Add a new mailbox to the list
function addMailbox() {
  var input = document.getElementById("newMailbox");
  var email = input.value.trim().toLowerCase();

  if (!email) {
    showStatus("Ingrese un correo para agregar.", "error");
    return;
  }

  // Basic email format check
  if (!email.includes("@") || !email.includes(".")) {
    showStatus("Formato de correo inválido.", "error");
    return;
  }

  var mailboxes = getMailboxes();

  if (mailboxes.includes(email)) {
    showStatus("Este buzón ya está en la lista.", "error");
    return;
  }

  mailboxes.push(email);
  saveMailboxes(mailboxes);

  // Auto-set as default if it is the first one added
  if (mailboxes.length === 1) {
    saveDefault(email);
  }

  input.value = "";
  renderList();
  showStatus("Buzón agregado: " + email, "success");
}

// Mark the selected mailbox as the default
function setDefault() {
  if (!selectedMailbox) {
    showStatus("Seleccione un buzón de la lista primero.", "error");
    return;
  }

  saveDefault(selectedMailbox);
  renderList();
  showStatus("Predeterminado: " + selectedMailbox, "success");
}

// Remove the selected mailbox from the list
function removeMailbox() {
  if (!selectedMailbox) {
    showStatus("Seleccione un buzón de la lista primero.", "error");
    return;
  }

  var removed  = selectedMailbox;
  var mailboxes = getMailboxes().filter(function (mb) { return mb !== removed; });
  saveMailboxes(mailboxes);

  // If the removed one was the default, move default to the first remaining one
  if (getDefault() === removed) {
    saveDefault(mailboxes.length > 0 ? mailboxes[0] : "");
  }

  selectedMailbox = null;
  renderList();
  showStatus("Buzón eliminado: " + removed, "success");
}

function showStatus(msg, type) {
  var el = document.getElementById("statusMsg");
  el.textContent = msg;
  el.className = "status status-" + type;
  el.style.display = "block";
  setTimeout(function () { el.style.display = "none"; }, 3500);
}

function goBack() {
  window.location.href = "taskpane.html";
}
