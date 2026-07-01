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

// Simulación completa de SpreadsheetApp conectada en tiempo real a Supabase
const SUPABASE_URL = process.env.SUPABASE_URL || 'https://kjcnotrxxthnzpgljeus.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtqY25vdHJ4eHRobnpwZ2xqZXVzIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MjY5ODY3MywiZXhwIjoyMDk4Mjc0NjczfQ.2yOpoM3C9ejhxA9hY0g88bkyU6KShhSaFHfnaBOLGiU';

function syncSupabase(table, method, body = null, query = '') {
  try {
    const { execFileSync } = require('child_process');
    const workerScript = `
      const https = require('https');
      const urlStr = process.argv[1];
      const method = process.argv[2];
      const headers = JSON.parse(process.argv[3]);
      const payload = process.argv[4];

      const url = new URL(urlStr);
      const req = https.request(url, { method: method, headers: headers }, res => {
        let d = ''; res.on('data', c => d += c);
        res.on('end', () => process.stdout.write(res.statusCode + '|||' + d));
      });
      if (payload) req.write(payload);
      req.end();
    `;
    const fullUrl = `${SUPABASE_URL}/rest/v1/${table}${query}`;
    const headers = {
      'apikey': SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
      'Content-Type': 'application/json',
      'Prefer': 'resolution=merge-duplicates'
    };
    const out = execFileSync(process.execPath, ['-e', workerScript, fullUrl, method, JSON.stringify(headers), body ? JSON.stringify(body) : ''], { encoding: 'utf8' });
    const parts = out.split('|||');
    if (parts[1]) return JSON.parse(parts[1]);
    return [];
  } catch(e) {
    console.error(`[SUPABASE ERROR ${table}]`, e.message);
    return [];
  }
}

const tableMap = {
  'Usuarios': 'usuarios',
  'Pagos': 'pagos',
  'Categorias': 'categorias',
  'Admins': 'admins',
  'Torneos': 'torneos',
  'Finanzas_Torneos': 'finanzas_torneos',
  'Partidos': 'partidos',
  'Logs_Audit': 'logs_audit'
};

const SHEET_HEADERS = {
  'Admins': ["Email"],
  'Categorias': ["Category_ID", "Name", "Coach", "Monthly_Fee", "Torneos"],
  'Usuarios': ["Email", "Role", "Name", "Phone", "Category", "DNI", "BirthDate", "Age", "JoinDate", "BloodType", "MedicalFit", "ObraSocial", "EmergencyContact", "EmergencyPhone", "ParentName", "ParentPhone", "Notes", "Photo", "Username", "Password"],
  'Pagos': ["Payment_ID", "Email", "Month", "Amount", "Status", "MP_Link", "Collected_By", "Collected_At"],
  'Torneos': ["Torneo_ID", "Name", "Category"],
  'Finanzas_Torneos': ["Movimiento_ID", "Torneo_ID", "Type", "Concept", "Amount", "Date", "Payment_Method"],
  'Partidos': ["Partido_ID", "Torneo_ID", "Date", "Opponent", "Location", "Result", "Scorers", "Cards", "MVP", "Summary"],
  'Logs_Audit': ["Timestamp", "User", "Action", "Details"]
};

const mockSpreadsheet = {
  getSheetByName: (sheetName) => {
    const sbTable = tableMap[sheetName] || sheetName.toLowerCase();
    return {
      getDataRange: () => ({
        getValues: () => {
          const headers = SHEET_HEADERS[sheetName] || [];
          // Consultar filas desde Supabase
          const sbRows = syncSupabase(sbTable, 'GET', null, '?select=*');
          if (Array.isArray(sbRows)) {
            // Ordenar consistentemente por ID para mantener orden determinista
            sbRows.sort((a, b) => (a.id || 0) - (b.id || 0));
            // Mapear objetos de Supabase de vuelta a arrays según las cabeceras fijas
            const rows = sbRows.map(obj => {
              return headers.map(h => {
                const key = h.toLowerCase();
                return obj[key] !== undefined && obj[key] !== null ? obj[key] : '';
              });
            });
            return [headers, ...rows];
          }
          return [headers];
        }
      }),
      appendRow: (rowArray) => {
        const headers = SHEET_HEADERS[sheetName] || [];
        // Insertar en Supabase
        const obj = {};
        headers.forEach((h, idx) => {
          obj[h.toLowerCase()] = rowArray[idx] !== undefined ? rowArray[idx] : '';
        });
        syncSupabase(sbTable, 'POST', [obj]);
        console.log(`[SUPABASE SUCCESS] Fila añadida dinámicamente en ${sbTable}:`, obj);
      },
      clear: () => {
        // Eliminar todos los registros de la tabla en Supabase
        syncSupabase(sbTable, 'DELETE', null, '?id=gt.0');
        console.log(`[SUPABASE SUCCESS] Limpiada tabla ${sbTable}`);
      },
      getLastRow: () => {
        const sbRows = syncSupabase(sbTable, 'GET', null, '?select=id');
        return Array.isArray(sbRows) ? sbRows.length + 1 : 1;
      },
      getRange: (row, col) => {
        return {
          setValue: (val) => {
            const headers = SHEET_HEADERS[sheetName] || [];
            const colName = headers[col - 1];
            if (!colName) return;
            const fieldName = colName.toLowerCase();

            // Buscar la fila por su posición
            const sbRows = syncSupabase(sbTable, 'GET', null, '?select=*');
            if (!Array.isArray(sbRows)) return;
            sbRows.sort((a, b) => (a.id || 0) - (b.id || 0));
            
            const rowIndex = row - 2;
            if (rowIndex >= 0 && rowIndex < sbRows.length) {
              const targetObj = sbRows[rowIndex];
              
              // Resolver el identificador primario de la tabla
              let idCol = 'id';
              if (sheetName === 'Usuarios' || sheetName === 'Admins') idCol = 'email';
              else if (sheetName === 'Pagos') idCol = 'payment_id';
              else if (sheetName === 'Categorias') idCol = 'category_id';
              else if (sheetName === 'Torneos') idCol = 'torneo_id';
              else if (sheetName === 'Finanzas_Torneos') idCol = 'movimiento_id';
              else if (sheetName === 'Partidos') idCol = 'partido_id';

              const idVal = targetObj[idCol];
              if (idVal !== undefined && idVal !== null) {
                const patchObj = {};
                patchObj[fieldName] = val;
                syncSupabase(sbTable, 'PATCH', patchObj, `?${idCol}=eq.${encodeURIComponent(idVal)}`);
                console.log(`[SUPABASE SUCCESS] Celda editada dinámicamente en ${sbTable}: ${fieldName}=${val} para ${idCol}=${idVal}`);
              }
            }
          },
          setNumberFormat: () => {}
        };
      },
      deleteRow: (rowIndex1Based) => {
        const sbRows = syncSupabase(sbTable, 'GET', null, '?select=*');
        if (!Array.isArray(sbRows)) return;
        sbRows.sort((a, b) => (a.id || 0) - (b.id || 0));

        const rowIndex = rowIndex1Based - 2;
        if (rowIndex >= 0 && rowIndex < sbRows.length) {
          const targetObj = sbRows[rowIndex];
          
          let idCol = 'id';
          if (sheetName === 'Usuarios' || sheetName === 'Admins') idCol = 'email';
          else if (sheetName === 'Pagos') idCol = 'payment_id';
          else if (sheetName === 'Categorias') idCol = 'category_id';
          else if (sheetName === 'Torneos') idCol = 'torneo_id';
          else if (sheetName === 'Finanzas_Torneos') idCol = 'movimiento_id';
          else if (sheetName === 'Partidos') idCol = 'partido_id';

          const idVal = targetObj[idCol];
          if (idVal !== undefined && idVal !== null) {
            syncSupabase(sbTable, 'DELETE', null, `?${idCol}=eq.${encodeURIComponent(idVal)}`);
            console.log(`[SUPABASE SUCCESS] Fila eliminada dinámicamente en ${sbTable} para ${idCol}=${idVal}`);
          }
        }
      }
    };
  },
  insertSheet: (name) => mockSpreadsheet.getSheetByName(name),
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
// ENDPOINT API /api/login
// ==========================================
app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  if (!username) return res.status(400).json({ success: false, message: 'Falta usuario.' });

  try {
    // Buscar el usuario en Supabase por username
    let query = '?select=*&username=eq.' + encodeURIComponent(username);
    let rows = syncSupabase('usuarios', 'GET', null, query);
    if (!Array.isArray(rows) || rows.length === 0) {
      // Intentar búsqueda por email
      query = '?select=*&email=eq.' + encodeURIComponent(username);
      rows = syncSupabase('usuarios', 'GET', null, query);
    }
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.json({ success: false, message: 'Usuario no encontrado en el sistema.' });
    }
    const user = rows[0];
    if (password && user.password && user.password !== String(password)) {
      return res.json({ success: false, message: 'Clave incorrecta.' });
    }
    return res.json({
      success: true,
      email: user.email,
      role: user.role || 'Deportista',
      name: user.name || '',
      username: user.username || username,
      photo: user.photo || ''
    });
  } catch (err) {
    console.error('[LOGIN ERROR]', err.message);
    res.status(500).json({ success: false, message: 'Error interno del servidor.' });
  }
});

// ==========================================
// ENDPOINT API /api/migrate-usuarios (DDL - solo usar una vez)
// ==========================================
app.get('/api/migrate-usuarios', async (req, res) => {
  const secret = req.query.secret;
  if (secret !== 'HaedoFutsal2026') return res.status(403).send('Forbidden');
  try {
    const https = require('https');
    const steps = [
      "ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS username TEXT;",
      "ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS password TEXT;",
      "UPDATE usuarios SET username = split_part(email, '@', 1) WHERE username IS NULL OR username = '';",
      "UPDATE usuarios SET password = '1234' WHERE password IS NULL OR password = '';",
      "UPDATE usuarios SET username = 'admin' WHERE email = 'admin@futsalhaedo.com';",
      "UPDATE usuarios SET username = 'gpazos' WHERE email = 'guillermopazos@gmail.com';",
      "UPDATE usuarios SET username = 'coach' WHERE email = 'coach@futsalhaedo.com';",
      "UPDATE usuarios SET username = 'socio' WHERE email = 'deportista@futsalhaedo.com';"
    ];
    const PROJECT_REF = 'kjcnotrxxthnzpgljeus';
    const results = [];
    for (const sql of steps) {
      const result = await new Promise((resolve) => {
        const body = JSON.stringify({ query: sql });
        const r = https.request({
          hostname: 'api.supabase.com',
          path: '/v1/projects/' + PROJECT_REF + '/database/query',
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
        }, (resp) => {
          let d = ''; resp.on('data', c => d += c);
          resp.on('end', () => resolve({ status: resp.statusCode, body: d }));
        });
        r.on('error', e => resolve({ status: 0, body: e.message }));
        r.write(body); r.end();
      });
      results.push({ sql: sql.substring(0, 60), status: result.status, body: result.body.substring(0, 100) });
    }
    res.json({ ok: true, results });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ==========================================
// ENDPOINT API /api/run (PROXEA CON CLIENTE)
// ==========================================
app.post('/api/run', (req, res) => {
  const { functionName, args } = req.body;
  console.log(`[API RUN] Ejecutando función del servidor: ${functionName}(${JSON.stringify(args)})`);
  
  try {
    const ctx = getSandboxContext();
    const func = ctx[functionName];
    if (typeof func !== 'function') {
      throw new Error(`La función de backend '${functionName}' no existe en el servidor.`);
    }
    
    const result = func.apply(null, args || []);

    // Sincronización en vivo con Supabase en operaciones de escritura
    const writeOps = ['registrarSocioNuevo', 'marcarPagoComoPagado', 'generarPagoMercadoPago', 'registrarPartidoNuevo', 'registrarMovimientoTorneo', 'editarMovimientoTorneo', 'eliminarMovimientoTorneo'];
    if (writeOps.includes(functionName)) {
      try {
        const https = require('https');
        const SUPABASE_URL = process.env.SUPABASE_URL || 'https://kjcnotrxxthnzpgljeus.supabase.co';
        const SUPABASE_KEY = process.env.SUPABASE_KEY || 'sb_publishable_1XDuAL5LGylnk6SUgG3JHQ_stWGkIQ-';
        
        if (functionName === 'registrarSocioNuevo' && args[0]) {
          const s = args[0];
          const payload = JSON.stringify([{
            email: (s.Email || '').toLowerCase().trim(), role: s.Role || 'Deportista', name: s.Name || '', phone: s.Phone || '',
            category: s.Category || '', dni: s.DNI || '', birthdate: s.BirthDate || '', bloodtype: s.BloodType || 'O+',
            medicalfit: s.MedicalFit || 'Apto Físico Vigente', obrasocial: s.ObraSocial || 'Particular', notes: s.Notes || '',
            username: s.Username || '', password: s.Password || ''
          }]);
          const req = https.request(new URL('/rest/v1/usuarios', SUPABASE_URL), {
            method: 'POST', headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json', 'Prefer': 'resolution=merge-duplicates' }
          }, res => console.log('[SUPABASE SYNC] Nuevo socio en Supabase:', res.statusCode));
          req.write(payload); req.end();
        } else if (functionName === 'marcarPagoComoPagado' && args[0]) {
          const pid = args[0];
          const collector = args[1] || 'Admin';
          const nowStr = new Date().toISOString().replace("T", " ").substring(0, 16);
          const payload = JSON.stringify({ status: 'Pagado', collected_by: collector, collected_at: nowStr });
          const req = https.request(new URL('/rest/v1/pagos?payment_id=eq.' + encodeURIComponent(pid), SUPABASE_URL), {
            method: 'PATCH', headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json' }
          }, res => console.log('[SUPABASE SYNC] Pago marcado en Supabase:', res.statusCode));
          req.write(payload); req.end();
        }
      } catch(syncErr) {
        console.error('[SUPABASE SYNC ERROR]', syncErr.message);
      }
    }

    res.json({ result });
    
  } catch (error) {
    console.error(`[API ERROR] Error ejecutando ${functionName}:`, error.message);
    res.status(500).json({ error: error.message || error.toString() });
  }
});

// ==========================================
// WEBHOOK / IPN MERCADO PAGO NOTIFICATIONS
// ==========================================
app.post('/api/mp-webhook', (req, res) => {
  console.log('[MERCADO PAGO WEBHOOK RECEIVED]', JSON.stringify(req.body));
  
  // Responder 200 OK de inmediato a Mercado Pago para confirmar recepción
  res.status(200).send('OK');

  const payload = req.body;
  const paymentId = payload.data && payload.data.id;
  const eventType = payload.type || payload.action;

  if (paymentId && (eventType === 'payment' || eventType === 'payment.created')) {
    try {
      const https = require('https');
      const MP_ACCESS_TOKEN = process.env.MP_ACCESS_TOKEN || "APP_USR-3796695277109598-062312-3b036573c099307abf18c869ea76c0f6-1868352613";
      
      const mpUrl = `https://api.mercadopago.com/v1/payments/${paymentId}`;
      const options = {
        headers: { 'Authorization': 'Bearer ' + MP_ACCESS_TOKEN }
      };

      https.get(mpUrl, options, (mpRes) => {
        let body = '';
        mpRes.on('data', chunk => body += chunk);
        mpRes.on('end', () => {
          try {
            const paymentInfo = JSON.parse(body);
            console.log('[MP PAYMENT INFO]', paymentInfo.id, paymentInfo.status, paymentInfo.operation_type);

            if (paymentInfo.status === 'approved') {
              const amount = paymentInfo.transaction_amount;
              const payerName = paymentInfo.description || (paymentInfo.payer && (paymentInfo.payer.first_name + ' ' + (paymentInfo.payer.last_name || '')));
              const payerEmail = paymentInfo.payer && paymentInfo.payer.email;
              const operationType = paymentInfo.operation_type;

              console.log(`[MP ACCREDITATION] Acreditando pago de $${amount} por ${payerName} (${payerEmail}) via ${operationType}`);

              // Caso A: Pago directo por preferencia/checkout (trae external_reference)
              const extRef = paymentInfo.external_reference;
              if (extRef && extRef.startsWith('PAG-')) {
                syncSupabase('pagos', 'PATCH', { status: 'Pagado', collected_by: 'MercadoPago', collected_at: new Date().toISOString().substring(0, 10) }, `?payment_id=eq.${encodeURIComponent(extRef)}`);
                console.log(`[MP SUCCESS] Acreditado pago de checkout: ${extRef}`);
                return;
              }

              // Caso B: Transferencia bancaria o dinero recibido (CVU/Alias)
              const users = syncSupabase('usuarios', 'GET', null, '?select=*');
              if (Array.isArray(users)) {
                let matchedUser = null;
                if (payerEmail) {
                  matchedUser = users.find(u => u.email && u.email.toLowerCase().trim() === payerEmail.toLowerCase().trim());
                }
                if (!matchedUser && payerName) {
                  const queryName = payerName.toLowerCase().trim();
                  matchedUser = users.find(u => {
                    const dbName = (u.name || '').toLowerCase().trim();
                    return dbName.includes(queryName) || queryName.includes(dbName);
                  });
                }

                if (matchedUser) {
                  console.log(`[MP MATCH] Transferencia emparejada con socio: ${matchedUser.name} (${matchedUser.email})`);
                  const pagosSocio = syncSupabase('pagos', 'GET', null, `?email=eq.${encodeURIComponent(matchedUser.email)}&status=eq.Pendiente&order=month.asc`);
                  if (Array.isArray(pagosSocio) && pagosSocio.length > 0) {
                    const oldestPayment = pagosSocio[0];
                    syncSupabase('pagos', 'PATCH', { status: 'Pagado', collected_by: 'Transferencia MP', collected_at: new Date().toISOString().substring(0, 10) }, `?payment_id=eq.${encodeURIComponent(oldestPayment.payment_id)}`);
                    console.log(`[MP SUCCESS] Acreditada cuota ${oldestPayment.payment_id} (${oldestPayment.month}) para ${matchedUser.name}`);
                  }
                }
              }
            }
          } catch(err) {
            console.error('[MP WEBHOOK PROCESSING ERROR]', err.message);
          }
        });
      }).on('error', (err) => {
        console.error('[MP GET ERROR]', err.message);
      });
    } catch(err) {
      console.error('[MP WEBHOOK EXCEPTION]', err.message);
    }
  }
});

app.get('/run-migration', async (req, res) => {
  const { Client } = require('pg');
  
  // Extraer el project reference dinámicamente del URL de Supabase configurado en el entorno
  const supabaseUrl = process.env.SUPABASE_URL || 'https://kjcnotrxxthnzpgljeus.supabase.co';
  const match = supabaseUrl.match(/https:\/\/([^.]+)\.supabase\.co/);
  const ref = match ? match[1] : 'kjcnotrxxthnzpgljeus';
  
  // Lista de posibles hosts a intentar (directo primero, luego poolers)
  const targets = [
    { host: `db.${ref}.supabase.co`, port: 5432, user: 'postgres' },
    { host: `aws-0-sa-east-1.pooler.supabase.com`, port: 6543, user: `postgres.${ref}` },
    { host: `aws-0-us-east-1.pooler.supabase.com`, port: 6543, user: `postgres.${ref}` },
    { host: `aws-0-us-east-2.pooler.supabase.com`, port: 6543, user: `postgres.${ref}` },
    { host: `aws-0-us-west-2.pooler.supabase.com`, port: 6543, user: `postgres.${ref}` }
  ];

  let lastError = null;
  const attempts = [];

  for (const t of targets) {
    console.log(`[MIGRATION TRY] Intentando conectar a ${t.host}:${t.port} con usuario ${t.user}...`);
    const client = new Client({
      host: t.host,
      port: t.port,
      database: 'postgres',
      user: t.user,
      password: 'HaedoFutsal.2026',
      ssl: { rejectUnauthorized: false },
      connectionTimeoutMillis: 4000
    });

    try {
      await client.connect();
      console.log(`[MIGRATION SUCCESS] Conexión exitosa a ${t.host}!`);
      await client.query(`
        ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS username TEXT UNIQUE;
        ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS password TEXT;
        UPDATE usuarios SET username = split_part(email, '@', 1) WHERE username IS NULL;
        UPDATE usuarios SET password = '1234' WHERE password IS NULL;
      `);
      await client.end();
      return res.send(`✅ Migración ejecutada con éxito en Supabase usando ${t.host}! (Project Ref: ${ref})`);
    } catch (err) {
      console.error(`[MIGRATION FAIL] Error en ${t.host}:`, err.message);
      lastError = err;
      attempts.push({ host: t.host, user: t.user, error: err.message });
    }
  }

  res.status(500).json({
    message: '❌ Error: Ningún host de conexión funcionó.',
    project_ref_detected: ref,
    supabase_url_env: supabaseUrl,
    last_error: lastError.message,
    attempts: attempts
  });
});

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
