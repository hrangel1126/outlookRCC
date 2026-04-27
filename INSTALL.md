# Instrucciones de Instalación — RCC Mail (New Outlook)

## Requisitos

- Node.js instalado → https://nodejs.org (LTS version)
- New Outlook (Windows) o Outlook en Web (outlook.com / Microsoft 365)
- Acceso a los buzones compartidos en tu cuenta de Outlook

---

## Paso 1 — Primera vez: instalar dependencias y certs

Abre una terminal (CMD o PowerShell) en esta carpeta:

```
cd C:\HR\RCCAPP\VSTORCC\RCCAddIn\outlooknew
```

Instala dependencias npm:
```
npm install
```

Genera los íconos placeholder (azules, reemplazar luego):
```
node setup.js
```

Instala los certificados HTTPS de desarrollo (solo una vez):
```
npm run install-certs
```
> Esto instala un certificado de Microsoft en tu Windows para que Outlook acepte `localhost`.
> Puede pedir confirmación — acepta.

---

## Paso 2 — Iniciar el servidor local

```
npm start
```

Deberías ver:
```
✅ RCC Mail add-in server running:
   https://localhost:3000/src/taskpane.html
```

Verifica que funciona abriendo `https://localhost:3000/src/taskpane.html` en Edge o Chrome.
Debes ver la página "RCC Mail" con los botones.

**Deja esta terminal abierta mientras uses el add-in.**

---

## Paso 3 — Cargar el add-in en Outlook (Sideload)

### En New Outlook (Windows):

1. Abre **New Outlook**
2. Haz clic en el ícono **Apps** en la barra lateral izquierda
3. Clic en **"Agregar aplicaciones"** o **"Add Apps"**
4. Busca el link **"My add-ins"** o **"Mis complementos"** abajo
5. Clic en **"Add a custom add-in"** → **"Add from file..."**
6. Selecciona: `C:\HR\RCCAPP\VSTORCC\RCCAddIn\outlooknew\manifest.xml`
7. Acepta la advertencia de publisher desconocido

### En Outlook Web (outlook.com / M365):

1. Abre **outlook.com** o tu Outlook Web
2. Abre cualquier correo (para activar el contexto de lectura)
3. Haz clic en los **tres puntos (...)** → **"Get Add-ins"** o **"Obtener complementos"**
4. Clic en **"My add-ins"** → **"Add a custom add-in"** → **"Add from URL..."**
5. Ingresa: `https://localhost:3000/manifest.xml`
   > ⚠️ Solo funciona si el servidor local está corriendo en tu máquina. Para compartir con otros usuarios usa GitHub Pages o M365 Admin Center (ver SESSION.md).

---

## Paso 4 — Usar el add-in

### Encontrar los botones:

1. Abre cualquier correo en Outlook
2. En la cinta superior (Home tab), busca el grupo **"RCC Mail"**
3. Si no lo ves, clic en el ícono **Apps** en la cinta → busca **RCC Mail** → clic **Pin**

### Primera configuración:

1. Clic en **"Configuracion"**
2. Escribe el correo del buzón compartido (ej: `soporte@empresa.com`)
3. Clic **"+ Agregar Buzón"**
4. Selecciónalo en la lista → clic **"★ Predeterminado"**
5. Clic **"← Volver"**

### Enviar correo:

1. Clic en **"Enviar Correo"**
2. El buzón predeterminado ya aparece seleccionado en "De"
3. Llena Para, CC (opcional), Asunto y Mensaje
4. Clic **"Enviar"** → Outlook abre la ventana de redacción con los campos llenados
5. ⚠️ Verifica el campo **"De"** en la ventana de Outlook y cámbialo si es necesario
6. Clic **"Enviar"** en la ventana de Outlook

---

## Desinstalar

### New Outlook / Outlook Web:
1. Abre **Apps** → **My add-ins**
2. Encuentra **RCC Mail** → clic en los tres puntos → **Remove**

### Detener el servidor:
```
Ctrl+C  en la terminal donde corre npm start
```

---

## Despliegue para múltiples usuarios (sin servidor local)

### Opción A — GitHub Pages (gratis, sin servidor)
1. Sube la carpeta `outlooknew/` a un repositorio GitHub
2. Activa GitHub Pages en el repositorio
3. En `manifest.xml`, reemplaza TODAS las ocurrencias de `https://localhost:3000` con tu URL de GitHub Pages
4. Comparte el `manifest.xml` con los usuarios — cada uno lo carga con "Add from file"

### Opción B — M365 Admin Center (despliegue automático, sin instalación de usuarios)
1. Ve a **portal.microsoft.com** → **Settings** → **Integrated Apps**
2. Sube `manifest.xml` como custom app
3. Asigna a los usuarios o grupos
4. El add-in aparece automáticamente en su Outlook sin que hagan nada

---

## Solución de problemas

| Problema | Causa probable | Solución |
|---|---|---|
| Botones no aparecen en cinta | Add-in en menú Apps | Clic Apps → Pin to ribbon |
| Error de certificado | Certs no instalados | `npm run install-certs` |
| "Not found" al cargar | Servidor no corre | `npm start` |
| Buzones no se guardan | localStorage bloqueado | Verifica que Outlook no bloquea cookies/storage |
| Ventana de correo no abre | Permisos insuficientes | El manifest necesita `ReadWriteMailbox` — verificar manifest.xml |
