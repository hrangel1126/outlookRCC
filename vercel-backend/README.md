# RCC Mail — Vercel Backend

API serverless que maneja autenticación con Microsoft y envío de correos via Graph API.
El add-in de Outlook solo manda los datos del correo — este backend tiene el token.

---

## Arquitectura

```
[Outlook Add-in]  →  POST /api/send-email  →  [Vercel Backend]  →  [Microsoft Graph API]
                       (x-api-key header)         (usa refresh token)      (envía correo)
```

El usuario hace login UNA VEZ en `/api/auth/login`. El refresh token queda guardado
en variables de entorno de Vercel y se renueva automáticamente en cada uso.

---

## Setup — paso a paso

### 1. Registrar app en Azure (necesario, 5 minutos)

Puedes usar CUALQUIER cuenta Microsoft, incluyendo una cuenta personal gratuita:

1. Ir a portal.azure.com (con cuenta outlook.com, hotmail.com, o cualquier M365)
2. **App registrations** → **New registration**
   - Name: `RCC Mail Backend`
   - Supported account types: **"Accounts in any organizational directory (Multitenant) and personal Microsoft accounts"**
   - Redirect URI: `Web` → `https://TU-APP.vercel.app/api/auth/callback`
   - Click **Register**
3. Copiar el **Application (client) ID** → es tu `CLIENT_ID`
4. **Certificates & secrets** → **New client secret** → copiar el valor → es tu `CLIENT_SECRET`
5. **API permissions** → **Add permission** → **Microsoft Graph** → **Delegated**:
   - `Mail.Send`
   - `Mail.Send.Shared`
   - `User.Read`
   - `offline_access`
   → Click **Grant admin consent** (si tienes permisos), o el usuario consiente al hacer login

### 2. Deploy en Vercel

```bash
# Instalar Vercel CLI
npm install -g vercel

# Entrar en esta carpeta
cd vercel-backend

# Deploy
vercel

# Tomar nota de la URL que Vercel asigna (ej: https://rcc-mail-xyz.vercel.app)
```

### 3. Configurar variables de entorno en Vercel

En el dashboard de Vercel → tu proyecto → **Settings → Environment Variables**:

| Variable | Valor |
|---|---|
| `CLIENT_ID` | El Application (client) ID de Azure |
| `CLIENT_SECRET` | El secret que creaste en Azure |
| `TENANT_ID` | `common` |
| `VERCEL_URL` | `https://tu-app.vercel.app` (sin slash final) |
| `API_KEY` | Cualquier string aleatorio largo (ej: `abc123xyz789...`) |
| `REFRESH_TOKEN` | Lo obtienes en el paso 4 |

### 4. Hacer login y obtener el refresh token

Una vez desplegado y con las variables configuradas (excepto REFRESH_TOKEN):

1. Abrir en el browser: `https://tu-app.vercel.app/api/auth/login`
2. Iniciar sesión con la cuenta Microsoft desde la que se enviarán los correos
3. La página mostrará el **REFRESH_TOKEN**
4. Copiar ese token → ir a Vercel → Settings → Environment Variables → agregar `REFRESH_TOKEN`
5. Hacer **redeploy** para que tome efecto

### 5. Actualizar el add-in

En `src/compose.js`, actualizar las dos líneas al inicio:

```js
var VERCEL_API = "https://tu-app.vercel.app"; // tu URL real de Vercel
var API_KEY    = "tu-api-key-aqui";            // el mismo API_KEY de Vercel env vars
```

Commit y push a GitHub — el add-in ya usará el backend.

---

## Probar el endpoint

```bash
curl -X POST https://tu-app.vercel.app/api/send-email \
  -H "Content-Type: application/json" \
  -H "x-api-key: tu-api-key" \
  -d '{
    "from": "buzon@empresa.com",
    "to": ["destinatario@empresa.com"],
    "subject": "Prueba RCC",
    "body": "Este correo fue enviado via Vercel + Graph API"
  }'
```

Respuesta esperada:
```json
{ "success": true, "message": "Correo enviado desde buzon@empresa.com" }
```

---

## Envío programático / bulk

Para enviar múltiples correos desde cualquier script:

```js
// Node.js example
const recipients = ["a@x.com", "b@x.com", "c@x.com"];

for (const email of recipients) {
    await fetch("https://tu-app.vercel.app/api/send-email", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "x-api-key": "tu-api-key"
        },
        body: JSON.stringify({
            from:    "buzon@empresa.com",
            to:      [email],
            subject: "Tu asunto",
            body:    "Tu mensaje"
        })
    });
}
```

No se necesita Outlook abierto. No se necesita autenticación en el script — el API key es suficiente.

---

## Archivos

| Archivo | Descripción |
|---|---|
| `api/auth/login.js` | Redirige al login de Microsoft |
| `api/auth/callback.js` | Recibe el token después del login, muestra el refresh token |
| `api/token.js` | Helper interno: obtiene access token usando el refresh token guardado |
| `api/send-email.js` | Endpoint principal: recibe datos del correo y lo envía via Graph API |
| `vercel.json` | Configuración de Vercel (Node.js 20) |
| `.env.example` | Plantilla de variables de entorno |
