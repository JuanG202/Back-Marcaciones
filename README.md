# Back-Marcaciones
Este proyecto es un sistema de registro de marcaciones (asistencia) desarrollado con HTML, JavaScript (Frontend) y Node.js (Backend). Los datos se guardan automáticamente en un archivo Excel en Google Drive mediante integración con Gmail/Google Sheets.

## Configuración

### Variables de Entorno
Crea un archivo `.env` en la raíz del proyecto con el siguiente contenido:
```
PORT=3000
EXCEL_PATH=marcaciones.xlsx
```

### Credenciales de Google Drive
1. Ve a [Google Cloud Console](https://console.cloud.google.com/)
2. Crea un nuevo proyecto o selecciona uno existente
3. Habilita la API de Google Drive
4. Crea credenciales de tipo "Cuenta de servicio"
5. Descarga el archivo JSON de credenciales
6. Renómbralo a `credentials.json` y colócalo en la raíz del proyecto

**IMPORTANTE**: No subas el archivo `credentials.json` a Git/GitHub. Este archivo contiene información sensible.

## Instalación
```bash
npm install
```

## Ejecución
```bash
node index.js
```

El servidor se iniciará en http://localhost:3000
