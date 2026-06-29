/**
 * Futsal Haedo ERP - Servidor de Desarrollo Local (Simulador Apps Script)
 * Ejecuta este servidor con Node.js para probar la aplicación al 100% en localhost.
 * 
 * Este script simula el motor de Google Apps Script y Google Sheets, leyendo/escribiendo
 * de forma persistente en un archivo local llamado 'db.json'.
 */

const express = require('express');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const os = require('os');
const app = express();
const PORT = process.env.PORT || 3000;
const DB_FILE = path.join(__dirname, 'db.json');

app.use(express.json());
app.use(express.static(__dirname));

// Servir el archivo estático Index.html
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'Index.html'));
});

// ==========================================
// SIMULACIÓN DE GOOGLE APPS SCRIPT (MOCKS GLOBALES)
// ==========================================

// Leer la base de datos JSON local
function readDb() {
  if (!fs.existsSync(DB_FILE)) {
    throw new Error("Base de datos local 'db.json' no encontrada. Por favor inicialízala.");
  }
  return JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
}

// Guardar en la base de datos JSON local
function writeDb(data) {
  fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// Simular el entorno global de Google Apps Script en Node.js
global.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: (key) => {
      if (key === "MP_ACCESS_TOKEN") {
        if (process.env.MP_ACCESS_TOKEN) return process.env.MP_ACCESS_TOKEN;
        try {
          const db = readDb();
          if (db.Config && db.Config.MP_ACCESS_TOKEN) return db.Config.MP_ACCESS_TOKEN;
        } catch(e) {}
        return "MOCK_MP_ACCESS_TOKEN_DEVELOPMENT";
      }
      return null;
    }
  })
};

global.HtmlService = {
  XFrameOptionsMode: { ALLOWALL: "ALLOWALL" }
};

// Simulación completa de SpreadsheetApp conectada a db.json
const mockSpreadsheet = {
  getSheetByName: (sheetName) => {
    return {
      getDataRange: () => ({
        getValues: () => {
          const db = readDb();
          const table = db[sheetName];
          if (!table) {
            console.error(`Error: No se encontró la tabla '${sheetName}' en db.json`);
            return [];
          }
          // Devolvemos las cabeceras + las filas
          return [table.headers, ...table.rows];
        }
      }),
      appendRow: (rowArray) => {
        const db = readDb();
        if (!db[sheetName]) db[sheetName] = { headers: [], rows: [] };
        db[sheetName].rows.push(rowArray);
        writeDb(db);
        console.log(`[DB SUCCESS] Fila añadida en ${sheetName}:`, rowArray);
      },
      clear: () => {
        // No es necesario en operaciones normales de desarrollo local, pero por compatibilidad
        const db = readDb();
        if (db[sheetName]) db[sheetName].rows = [];
        writeDb(db);
      },
      getLastRow: () => {
        const db = readDb();
        return db[sheetName] ? db[sheetName].rows.length + 1 : 1;
      },
      getRange: (row, col) => {
        return {
          setValue: (val) => {
            const db = readDb();
            const table = db[sheetName];
            if (!table) return;
            
            // row y col están basados en 1. Las cabeceras son row = 1.
            const rowIndex = row - 2; // Índice en el array de filas (rows)
            const colIndex = col - 1; // Índice en el array de columnas
            
            if (rowIndex >= 0 && rowIndex < table.rows.length) {
              table.rows[rowIndex][colIndex] = val;
              writeDb(db);
              console.log(`[DB SUCCESS] Celda en ${sheetName} modificada (Fila ${row}, Col ${col}) a:`, val);
            }
          },
          setNumberFormat: () => {} // Simulación vacía por compatibilidad
        };
      },
      deleteRow: (rowIndex1Based) => {
        const db = readDb();
        const table = db[sheetName];
        if (table && table.rows) {
          const idx = rowIndex1Based - 2; // header is row 1
          if (idx >= 0 && idx < table.rows.length) {
            const removed = table.rows.splice(idx, 1);
            writeDb(db);
            console.log(`[DB SUCCESS] Fila ${rowIndex1Based} eliminada en ${sheetName}:`, removed);
          }
        }
      }
    };
  },
  insertSheet: (name) => {
    // Por si se vuelve a correr el inicializador localmente
    const db = readDb();
    if (!db[name]) {
      db[name] = { headers: [], rows: [] };
      writeDb(db);
    }
    return mockSpreadsheet.getSheetByName(name);
  },
  deleteSheet: () => {},
  getSheets: () => [{}, {}]
};

global.SpreadsheetApp = {
  getActiveSpreadsheet: () => mockSpreadsheet,
  openById: () => mockSpreadsheet,
  flush: () => {}
};

// Simulación de UrlFetchApp que mockea la llamada a Mercado Pago
global.UrlFetchApp = {
  fetch: (url, options) => {
    console.log(`[EXTERNAL API CALL] Petición a: ${url}`);
    const method = (options && options.method) ? options.method.toUpperCase() : 'GET';
    const headers = (options && options.headers) || {};
    const payload = (options && options.payload) ? options.payload : '';

    const workerScript = `
      const https = require('https');
      const http = require('http');
      const urlStr = process.argv[1];
      const method = process.argv[2];
      const headers = JSON.parse(process.argv[3]);
      const payload = process.argv[4];

      const parsedUrl = new URL(urlStr);
      const client = parsedUrl.protocol === 'https:' ? https : http;

      const req = client.request(parsedUrl, {
        method: method,
        headers: headers
      }, res => {
        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', () => {
          process.stdout.write(res.statusCode + '|||' + data);
        });
      });
      if (payload) req.write(payload);
      req.end();
    `;

    try {
      const { execFileSync } = require('child_process');
      const out = execFileSync(process.execPath, ['-e', workerScript, url, method, JSON.stringify(headers), payload], { encoding: 'utf8' });
      const parts = out.split('|||');
      const statusCode = parseInt(parts[0] || '200');
      const contentText = parts.slice(1).join('|||');
      return {
        getResponseCode: () => statusCode,
        getContentText: () => contentText
      };
    } catch (err) {
      console.error('[EXTERNAL API ERROR]:', err.message);
      throw err;
    }
  }
};

function getSandboxContext() {
  const backendCodePath = path.join(__dirname, 'Codigo.js');
  const ctx = {
    SpreadsheetApp: global.SpreadsheetApp,
    PropertiesService: global.PropertiesService,
    UrlFetchApp: global.UrlFetchApp,
    HtmlService: global.HtmlService,
    console: console,
    parseFloat: parseFloat,
    parseInt: parseInt,
    Date: Date,
    Error: Error,
    JSON: JSON,
    Math: Math,
    SPREADSHEET_ID: "1nDZWIxdGPWK8YY86JcOW9YCwTuApQX380H7Wk4JGjWg",
    HOJA_PAGOS: "Pagos",
    HOJA_USUARIOS: "Usuarios",
    HOJA_CATEGORIAS: "Categorias",
    HOJA_ADMINS: "Admins",
    HOJA_TORNEOS: "Torneos",
    HOJA_FINANZAS: "Finanzas_Torneos",
    HOJA_PARTIDOS: "Partidos",
    HOJA_LOGS: "Logs_Audit"
  };
  vm.createContext(ctx);
  if (fs.existsSync(backendCodePath)) {
    const code = fs.readFileSync(backendCodePath, 'utf8');
    vm.runInContext(code, ctx);
  }
  return ctx;
}

console.log("¡Backend de Codigo.js cargado y simulado con éxito!");

// ==========================================
// ENDPOINT API /api/run (PROXEA CON CLIENTE)
// ==========================================
app.post('/api/run', (req, res) => {
  const { functionName, args } = req.body;
  console.log(`[API RUN] Ejecutando función del servidor: ${functionName}(${JSON.stringify(args)})`);
  
  try {
    const ctx = getSandboxContext();
    const func = ctx[functionName];
    // Comprobar si la función existe en el contexto sandbox
    if (typeof func !== 'function') {
      throw new Error(`La función de backend '${functionName}' no existe en el servidor.`);
    }
    
    // Ejecutar la función con los argumentos
    const result = func.apply(null, args || []);
    res.json({ result });
    
  } catch (error) {
    console.error(`[API ERROR] Error ejecutando ${functionName}:`, error.message);
    res.status(500).json({ error: error.message || error.toString() });
  }
});

// Levantar el servidor en todas las interfaces de red (0.0.0.0)
app.listen(PORT, '0.0.0.0', () => {
  const interfaces = os.networkInterfaces();
  const addresses = [];
  for (const k in interfaces) {
    for (const k2 of interfaces[k]) {
      if (k2.family === 'IPv4' && !k2.internal) {
        addresses.push(k2.address);
      }
    }
  }
  
  console.log(`\n======================================================`);
  console.log(`⚽ ERP Futsal Haedo corriendo!`);
  console.log(`💻 En esta PC: http://localhost:${PORT}`);
  addresses.forEach(ip => {
    console.log(`📱 En tu Red Wi-Fi (celulares/tablets): http://${ip}:${PORT}`);
  });
  console.log(`======================================================`);
  console.log(`Modificaciones y pruebas activas y persistentes.`);
  console.log(`Cualquier cambio se guardará en tu archivo local 'db.json'.`);
  console.log(`======================================================\n`);
});
