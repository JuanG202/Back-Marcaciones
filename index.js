const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

// =============================================
// Configuración inicial y verificación
// =============================================

console.log('Iniciando servidor de marcaciones...');
console.log('Verificando configuración...');

// Verificar variables de entorno esenciales
const REQUIRED_ENV = ['CARPETA_DRIVE_ID', 'CREDENTIALS_PATH'];
REQUIRED_ENV.forEach(env => {
  if (!process.env[env]) {
    console.error(`ERROR: Variable de entorno requerida faltante: ${env}`);
    process.exit(1);
  }
});

// Verificar archivo de credenciales
if (!fs.existsSync(path.join(__dirname, process.env.CREDENTIALS_PATH))) {
  console.error('ERROR: No se encuentra el archivo credentials.json');
  process.exit(1);
}

// =============================================
// Configuración de Express y Middlewares
// =============================================

const app = express();

// Configuración de CORS
const allowedOrigins = [
  'https://registro-marcaciones.vercel.app',
  'http://localhost:3000', // Desarrollo frontend local
  'https://localhost:3000' // Para desarrollo con HTTPS
];

const corsOptions = {
  origin: function (origin, callback) {
    if (!origin || allowedOrigins.includes(origin)) {
      callback(null, true);
    } else {
      callback(new Error('Origen no permitido por CORS'));
    }
  },
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
  optionsSuccessStatus: 200
};

app.use(cors(corsOptions));
app.options('*', cors(corsOptions)); // Habilitar preflight para todas las rutas

// Middlewares para parsear el cuerpo de las peticiones
app.use(bodyParser.json({ limit: '10mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '10mb' }));

// Middleware de logging
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.ip} ${req.method} ${req.path}`);
  next();
});

// =============================================
// Configuración de rutas de archivos
// =============================================

const EXCEL_PATH = path.join(__dirname, process.env.EXCEL_PATH || 'marcaciones.xlsx');
const CARPETA_DRIVE_ID = process.env.CARPETA_DRIVE_ID;
const CREDENTIALS_PATH = path.join(__dirname, process.env.CREDENTIALS_PATH);

// =============================================
// Funciones auxiliares
// =============================================

/**
 * Guarda los datos en un archivo Excel local
 * @param {Object} datos - Datos a guardar
 */
function guardarEnExcel(datos) {
  console.log('Iniciando guardado en Excel...');
  
  let registros = [];
  let workbook;

  try {
    if (fs.existsSync(EXCEL_PATH)) {
      workbook = XLSX.readFile(EXCEL_PATH);
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      registros = XLSX.utils.sheet_to_json(hoja);
    } else {
      workbook = XLSX.utils.book_new();
    }

    registros.push(datos);
    const nuevaHoja = XLSX.utils.json_to_sheet(registros);
    
    if (workbook.SheetNames.length > 0) {
      workbook.Sheets[workbook.SheetNames[0]] = nuevaHoja;
    } else {
      XLSX.utils.book_append_sheet(workbook, nuevaHoja, 'Marcaciones');
    }
    
    XLSX.writeFile(workbook, EXCEL_PATH);
    console.log('Excel guardado exitosamente');
  } catch (error) {
    console.error('Error al guardar en Excel:', error);
    throw error;
  }
}

/**
 * Sube el archivo Excel a Google Drive
 */
async function subirArchivoAGoogleDrive() {
  if (!fs.existsSync(EXCEL_PATH)) {
    throw new Error('El archivo Excel no existe localmente');
  }

  console.log('Autenticando con Google Drive...');
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS_PATH,
    scopes: ['https://www.googleapis.com/auth/drive.file']
  });

  const drive = google.drive({ version: 'v3', auth });

  try {
    console.log('Buscando archivo existente en Drive...');
    const respuesta = await drive.files.list({
      q: `name='marcaciones.xlsx' and '${CARPETA_DRIVE_ID}' in parents`,
      fields: 'files(id, name)'
    });

    let archivoId;
    if (respuesta.data.files.length > 0) {
      archivoId = respuesta.data.files[0].id;
      console.log('Archivo existente encontrado, actualizando...');
    } else {
      console.log('Creando nuevo archivo en Drive...');
      const fileMetadata = {
        name: 'marcaciones.xlsx',
        parents: [CARPETA_DRIVE_ID]
      };
      
      const media = {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: fs.createReadStream(EXCEL_PATH)
      };

      const archivo = await drive.files.create({
        resource: fileMetadata,
        media,
        fields: 'id, webViewLink'
      });

      archivoId = archivo.data.id;
    }

    // Actualizar el archivo
    const media = {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: fs.createReadStream(EXCEL_PATH)
    };

    const archivoActualizado = await drive.files.update({
      fileId: archivoId,
      media,
      fields: 'id, webViewLink'
    });

    console.log('Archivo actualizado en Drive:', archivoActualizado.data.webViewLink);
    return {
      mensaje: 'Registro guardado en Google Drive',
      archivoUrl: archivoActualizado.data.webViewLink
    };
  } catch (error) {
    console.error('Error al interactuar con Google Drive:', error);
    throw error;
  }
}

/**
 * Combina una fecha con la hora actual
 */
function combinarFechaConHoraActual(fecha) {
  if (!fecha) return '';
  
  const ahora = new Date();
  const fechaObj = new Date(fecha);
  
  fechaObj.setHours(ahora.getHours());
  fechaObj.setMinutes(ahora.getMinutes());
  fechaObj.setSeconds(ahora.getSeconds());

  return fechaObj.toLocaleString('es-ES', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
}

// =============================================
// Endpoints de la API
// =============================================

/**
 * Endpoint para registrar una nueva marcación
 */
app.post('/registrar', async (req, res) => {
  try {
    console.log('Datos recibidos:', req.body);
    
    const { nombre, cedula, agencia, horaEntrada, horaSalida, observaciones } = req.body;

    // Validación de campos obligatorios
    if (!nombre || !cedula || !agencia) {
      return res.status(400).json({ 
        success: false, 
        message: 'Nombre, cédula y agencia son campos obligatorios' 
      });
    }

    // Formatear fechas
    const fechaHoraEntrada = horaEntrada ? combinarFechaConHoraActual(horaEntrada) : '';
    const fechaHoraSalida = horaSalida ? combinarFechaConHoraActual(horaSalida) : '';

    const registro = {
      nombre,
      cedula,
      agencia,
      horaEntrada: fechaHoraEntrada,
      horaSalida: fechaHoraSalida,
      observaciones: observaciones || '',
      fechaRegistro: new Date().toISOString(),
      ip: req.ip
    };

    // Guardar en Excel local
    guardarEnExcel(registro);
    
    // Subir a Google Drive
    const driveInfo = await subirArchivoAGoogleDrive();
    
    res.json({ 
      success: true,
      message: 'Registro exitoso',
      data: registro,
      driveInfo
    });
  } catch (error) {
    console.error('Error en endpoint /registrar:', error);
    res.status(500).json({ 
      success: false,
      message: 'Error al procesar el registro',
      error: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
});

/**
 * Endpoint para obtener todos los registros
 */
app.get('/registros', async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.json({ success: true, data: [] });
    }

    const workbook = XLSX.readFile(EXCEL_PATH);
    const hoja = workbook.Sheets[workbook.SheetNames[0]];
    const registros = XLSX.utils.sheet_to_json(hoja);

    res.json({ success: true, data: registros });
  } catch (error) {
    console.error('Error en endpoint /registros:', error);
    res.status(500).json({ 
      success: false, 
      message: 'Error al obtener registros',
      error: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
});

/**
 * Endpoint de verificación de salud del servidor
 */
app.get('/health', (req, res) => {
  const status = {
    status: 'OK',
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    memoryUsage: process.memoryUsage(),
    environment: process.env.NODE_ENV || 'development',
    excelFileExists: fs.existsSync(EXCEL_PATH),
    driveConfig: CARPETA_DRIVE_ID ? 'Configurado' : 'No configurado'
  };

  res.status(200).json(status);
});

// =============================================
// Manejo de errores global
// =============================================

app.use((err, req, res, next) => {
  console.error('Error global:', err);
  res.status(500).json({
    success: false,
    message: 'Error interno del servidor',
    error: process.env.NODE_ENV === 'development' ? err.message : undefined
  });
});

// =============================================
// Iniciar el servidor
// =============================================

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor escuchando en puerto ${PORT}`);
  console.log('Orígenes permitidos:', allowedOrigins);
  console.log('Ruta archivo Excel:', EXCEL_PATH);
  console.log('ID Carpeta Drive:', CARPETA_DRIVE_ID);
});