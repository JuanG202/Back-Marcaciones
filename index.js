const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

// Verificación inicial de configuración
console.log('Verificando configuración...');
if (!process.env.CARPETA_DRIVE_ID) {
  console.error('Error: CARPETA_DRIVE_ID no está definido en el archivo .env');
}
if (!fs.existsSync(process.env.CREDENTIALS_PATH || 'credentials.json')) {
  console.error('Error: No se encuentra el archivo credentials.json');
}

const app = express();

// Configuración básica de CORS
app.use(cors());

// Middleware para asegurar los headers CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

app.use(bodyParser.json());

// Configuración de archivos
const EXCEL_PATH = path.join(__dirname, process.env.EXCEL_PATH || 'marcaciones.xlsx');
const CARPETA_DRIVE_ID = process.env.CARPETA_DRIVE_ID;
const CREDENTIALS_PATH = path.join(__dirname, process.env.CREDENTIALS_PATH || 'credentials.json');

// Función para guardar en Excel local
function guardarEnExcel(datos) {
  console.log('Iniciando guardado en Excel...');
  console.log('Ruta del archivo Excel:', EXCEL_PATH);
  
  let registros = [];

  try {
    if (fs.existsSync(EXCEL_PATH)) {
      console.log('Archivo Excel existente encontrado, leyendo datos...');
      const workbook = XLSX.readFile(EXCEL_PATH);
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      registros = XLSX.utils.sheet_to_json(hoja);
      console.log(`Registros existentes cargados: ${registros.length}`);
    } else {
      console.log('No existe archivo Excel previo, se creará uno nuevo');
    }

    registros.push(datos);
    console.log('Nuevo registro agregado');

    const nuevaHoja = XLSX.utils.json_to_sheet(registros);
    const nuevoLibro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'Marcaciones');
    
    console.log('Guardando archivo Excel...');
    XLSX.writeFile(nuevoLibro, EXCEL_PATH);
    console.log('Archivo Excel guardado exitosamente');
    
    if (!fs.existsSync(EXCEL_PATH)) {
      throw new Error('El archivo no se guardó correctamente');
    }
    
    const stats = fs.statSync(EXCEL_PATH);
    console.log(`Tamaño del archivo Excel: ${stats.size} bytes`);
    
  } catch (error) {
    console.error('Error al guardar en Excel:', error);
    throw new Error(`Error al guardar en Excel: ${error.message}`);
  }
}

// Función para verificar si la carpeta existe
async function verificarCarpeta(drive, carpetaId) {
  try {
    const response = await drive.files.get({
      fileId: carpetaId,
      fields: 'id, name'
    });
    return response.data;
  } catch (error) {
    console.error('Error al verificar la carpeta:', error);
    throw new Error(`No se puede acceder a la carpeta con ID ${carpetaId}. Verifica que existe y que tienes permisos.`);
  }
}

// Función para crear o encontrar carpeta en Drive
async function crearOEncontrarCarpeta(drive) {
  console.log('Buscando carpeta "Marcaciones" en Drive...');
  
  const busqueda = await drive.files.list({
    q: "mimeType='application/vnd.google-apps.folder' and name='Marcaciones' and trashed=false",
    fields: 'files(id, name, webViewLink)',
    spaces: 'drive'
  });

  if (busqueda.data.files.length > 0) {
    const carpeta = busqueda.data.files[0];
    console.log('Carpeta encontrada:', carpeta.name, 'ID:', carpeta.id);
    return carpeta;
  }

  console.log('Creando nueva carpeta "Marcaciones"...');
  const fileMetadata = {
    name: 'Marcaciones',
    mimeType: 'application/vnd.google-apps.folder'
  };

  const carpeta = await drive.files.create({
    resource: fileMetadata,
    fields: 'id, name, webViewLink'
  });

  await drive.permissions.create({
    fileId: carpeta.data.id,
    requestBody: {
      role: 'writer',
      type: 'anyone'
    }
  });

  console.log('Nueva carpeta creada con ID:', carpeta.data.id);
  return carpeta.data;
}

// Función para subir archivo a Google Drive
async function subirArchivoAGoogleDrive() {
  if (!fs.existsSync(EXCEL_PATH)) {
    throw new Error('El archivo Excel no existe');
  }

  console.log('Iniciando proceso de subida a Drive...');
  
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDENTIALS_PATH,
    scopes: ['https://www.googleapis.com/auth/drive']
  });

  const drive = google.drive({ version: 'v3', auth });
  
  // ID del archivo Excel en Drive (reemplaza con el correcto)
  const EXCEL_DRIVE_ID = '1mNYuHeBH0ODc4m8ajDTdjPBqkDDG7_hR';
  
  console.log('Actualizando archivo Excel en Drive...');
  const media = {
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    body: fs.createReadStream(EXCEL_PATH)
  };

  try {
    const actualizado = await drive.files.update({
      fileId: EXCEL_DRIVE_ID,
      media,
      fields: 'id, webViewLink'
    });

    console.log('Archivo actualizado exitosamente');
    return {
      archivoUrl: actualizado.data.webViewLink,
      mensaje: 'Registro guardado exitosamente en el archivo Excel de Google Drive'
    };
  } catch (error) {
    console.error('Error al actualizar archivo:', error);
    throw new Error('No se pudo actualizar el archivo en Drive. Verifica que el ID sea correcto y tengas permisos de edición.');
  }
}

// Función para combinar fecha con hora actual
function combinarFechaConHoraActual(fecha) {
  if (!fecha) return '';
  
  const ahora = new Date();
  const fechaSeleccionada = new Date(fecha);
  
  fechaSeleccionada.setHours(ahora.getHours());
  fechaSeleccionada.setMinutes(ahora.getMinutes());
  fechaSeleccionada.setSeconds(ahora.getSeconds());

  return fechaSeleccionada.toLocaleString('es-ES', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
}

// Endpoint para registrar
app.post('/registrar', async (req, res) => {
  try {
    console.log('=== Inicio de /registrar ===');
    console.log('Datos recibidos:', req.body);
    
    const { nombre, cedula, agencia, horaEntrada, horaSalida, observaciones } = req.body;

    if (!nombre || !cedula || !agencia) {
      return res.status(400).json({ 
        success: false, 
        message: 'Faltan datos obligatorios: nombre, cédula y agencia son requeridos' 
      });
    }

    const fechaHoraEntrada = horaEntrada ? combinarFechaConHoraActual(horaEntrada) : '';
    const fechaHoraSalida = horaSalida ? combinarFechaConHoraActual(horaSalida) : '';

    const datos = {
      nombre,
      cedula,
      agencia,
      horaEntrada: fechaHoraEntrada,
      horaSalida: fechaHoraSalida,
      observaciones: observaciones || ''
    };

    console.log('Guardando registro:', datos);
    guardarEnExcel(datos);
    const driveInfo = await subirArchivoAGoogleDrive();
    
    res.json({ 
      success: true,
      message: driveInfo.mensaje,
      data: datos,
      urls: {
        archivo: driveInfo.archivoUrl
      }
    });
  } catch (error) {
    console.error('=== Error en /registrar ===');
    console.error('Mensaje:', error.message);
    console.error('Stack:', error.stack);
    res.status(500).json({ 
      success: false,
      message: 'Error al guardar el registro: ' + error.message,
      error: error.stack
    });
  }
});

// Endpoint para obtener registros
app.get('/registros', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.json({ success: true, data: [] });
    }

    const workbook = XLSX.readFile(EXCEL_PATH);
    const hoja = workbook.Sheets[workbook.SheetNames[0]];
    const registros = XLSX.utils.sheet_to_json(hoja);

    res.json({ success: true, data: registros });
  } catch (error) {
    console.error('Error al obtener registros:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Error al obtener los registros' 
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
