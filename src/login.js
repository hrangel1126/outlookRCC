// login.js - Standalone MSAL login page

var msalLoaded = false;

function msalLoaded() {
    msalLoaded = true;
    showStatus("MSAL cargado", "info");
}

function msalFailed() {
    showStatus("Error: No se pudo cargar MSAL. Intenta recargar la página.", "error");
}

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

var GRAPH_SCOPES = ["Mail.Send", "Mail.Send.Shared", "User.Read"];

// Check for existing token on load
window.addEventListener("load", function() {
    // Wait a bit for MSAL to load
    setTimeout(function() {
        var token = localStorage.getItem("rcc_graph_token");
        if (token) {
            showToken(token);
        }
        
        if (typeof msal !== "undefined") {
            msalLoaded = true;
        }
    }, 2000);
});

async function login() {
    if (!msalLoaded && typeof msal === "undefined") {
        showStatus("Cargando MSAL, espera un momento...", "info");
        
        // Wait up to 5 seconds for MSAL
        for (var i = 0; i < 10; i++) {
            await new Promise(function(resolve) { setTimeout(resolve, 500); });
            if (typeof msal !== "undefined") {
                msalLoaded = true;
                break;
            }
        }
    }
    
    if (typeof msal === "undefined") {
        showStatus("Error: MSAL no se cargó. Recarga la página.", "error");
        return;
    }
    
    try {
        showStatus("Iniciando sesión...", "info");
        
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
        
        // Store token for add-in to use
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