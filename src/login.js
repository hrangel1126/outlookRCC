// login.js - Standalone MSAL login page with forced consent

var MSAL_CONFIG = {
    auth: {
        clientId: "3b387ca6-bd43-4396-b557-bbd5786405db",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.href
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
    }
};

// Force consent by adding prompt parameter
var LOGIN_REQUEST = {
    scopes: ["Mail.Send", "Mail.Send.Shared", "User.Read"],
    prompt: "consent"  // Forces consent screen to appear
};

function waitForMsal(callback) {
    if (typeof msal !== "undefined") {
        callback();
        return;
    }
    
    var attempts = 0;
    var interval = setInterval(function() {
        attempts++;
        if (typeof msal !== "undefined") {
            clearInterval(interval);
            callback();
        } else if (attempts > 20) {
            clearInterval(interval);
            showStatus("Error: No se pudo cargar la biblioteca de autenticación", "error");
        }
    }, 500);
}

window.addEventListener("load", function() {
    waitForMsal(function() {
        document.getElementById("loading").style.display = "none";
        document.getElementById("loginInstructions").style.display = "block";
        document.getElementById("loginBtn").style.display = "inline-block";
        document.getElementById("logoutBtn").style.display = "inline-block";
        
        var token = localStorage.getItem("rcc_graph_token");
        if (token) {
            showToken(token);
        }
    });
});

async function login() {
    try {
        showStatus("Iniciando sesión...", "info");
        
        var pca = new msal.PublicClientApplication(MSAL_CONFIG);
        
        // Clear cache to force fresh consent prompt
        pca.clearCache();
        
        // Use prompt=consent to force consent screen
        var result = await pca.acquireTokenPopup(LOGIN_REQUEST);
        
        localStorage.setItem("rcc_graph_token", result.accessToken);
        
        showToken(result.accessToken);
        showStatus("✓ Sesión iniciada correctamente!", "success");
        
    } catch (err) {
        showStatus("Error: " + err.message, "error");
    }
}

function logout() {
    localStorage.removeItem("rcc_graph_token");
    document.getElementById("tokenBox").className = "token-box";
    document.getElementById("tokenBox").textContent = "";
    showStatus("Sesión cerrada", "info");
}

function showToken(token) {
    var box = document.getElementById("tokenBox");
    var payload = decodeJwt(token);
    var expTime = payload && payload.exp ? new Date(payload.exp * 1000).toLocaleString() : "?";
    var user = payload ? (payload.preferred_username || payload.upn || "?") : "?";
    
    box.innerHTML = 
        "<strong>Token obtenido</strong><br/>" +
        "Usuario: " + user + "<br/>" +
        "Expira: " + expTime + "<br/><br/>" +
        "Token (primeros 50 caracteres):<br/>" +
        token.substring(0, 50) + "...";
    box.className = "token-box show";
}

function showStatus(msg, type) {
    var el = document.getElementById("status");
    el.textContent = msg;
    el.className = "status status-" + type;
    el.style.display = "block";
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