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

// Configurar CORS para aceptar solo peticiones de tu frontend en Vercel
const corsOptions = {
  origin: 'https://registro-marcaciones.vercel.app', // Cambia si tu frontend cambia
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
  optionsSuccessStatus: 204
};

app.use(cors(corsOptions));

// Middleware para responder a las peticiones OPTIONS (preflight)
app.options('*', cors(corsOptions));

// Body parser
app.use(bodyParser.json());

// --- Aquí va el resto de tu código ---

// Configuración de archivos
const EXCEL_PATH = path.join(__dirname, process.env.EXCEL_PATH || 'marcaciones.xlsx');
const CARPETA_DRIVE_ID = process.env.CARPETA_DRIVE_ID;
const CREDENTIALS_PATH = path.join(__dirname, process.env.CREDENTIALS_PATH || 'credentials.json');

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

// Funciones verificarCarpeta, crearOEncontrarCarpeta, subirArchivoAGoogleDrive, combinarFechaConHoraActual aquí igual (sin cambios)...

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
