const express = require('express');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const axios = require('axios');
const Tesseract = require('tesseract.js');

const auth = require('./auth');
const apiFuncs = require('./api');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));

// Serve frontend
app.get('/sw.js', (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache');
  res.setHeader('Content-Type', 'application/javascript');
  res.sendFile(path.join(__dirname, '../sw.js'));
});
app.get('/', (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache');
  res.sendFile(path.join(__dirname, '../Index.html'));
});
app.get('/Index.html', (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache');
  res.sendFile(path.join(__dirname, '../Index.html'));
});
app.use(express.static(path.join(__dirname, '..')));

// =====================================
// AUTH ROUTES
// =====================================
app.post('/api/login', auth.loginHandler);

// =====================================
// API RUN ROUTE
// =====================================
// Esta ruta reemplaza google.script.run
app.post('/api/run', auth.authenticateToken, async (req, res) => {
  const { functionName, args } = req.body;
  console.log(`[API RUN] ${functionName}`);
  
  try {
    // 1. Intentar correr la version nativa Node.js (rapida)
    if (apiFuncs[functionName]) {
      console.log(`[NATIVE] Ejecutando ${functionName} nativamente en Node.js`);
      const result = await apiFuncs[functionName](...(args || []));
      return res.json({ success: true, result });
    }
    
    // 2. Fallback a la version antigua de Codigo.js en la maquina virtual
    console.log(`[FALLBACK VM] Ejecutando ${functionName} en VM heredada`);
    const ctx = getSandboxContext();
    const func = ctx[functionName];
    if (typeof func !== 'function') {
      return res.status(404).json({ success: false, message: 'Funcion no encontrada' });
    }
    const result = func.apply(null, args || []);
    res.json({ success: true, result });
    
  } catch (err) {
    console.error(`[API RUN ERROR] ${functionName}:`, err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// =====================================
// MOCK VM PARA CODIGO.JS HEREDADO
// =====================================
function getSandboxContext() {
  const backendCodePath = path.join(__dirname, '../Codigo.js');
  let backendCode = '';
  if (fs.existsSync(backendCodePath)) backendCode = fs.readFileSync(backendCodePath, 'utf8');
  
  const ctx = vm.createContext({
    console: console,
    Math: Math,
    Date: Date,
    JSON: JSON,
    parseInt: parseInt,
    parseFloat: parseFloat,
    isNaN: isNaN,
    SpreadsheetApp: getMockSpreadsheetApp(),
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => process.env[key] || "MOCK_KEY"
      })
    },
    UrlFetchApp: {
      fetch: (url, options) => {
        // Ignorado, o implementar si es necesario para webhooks viejos
        return { getResponseCode: () => 200, getContentText: () => '{}' };
      }
    },
    Session: { getActiveUser: () => ({ getEmail: () => 'admin@futsalhaedo.com' }) },
    Utilities: { getUuid: () => require('crypto').randomUUID() },
    MailApp: { sendEmail: () => console.log('Mock email enviado') },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) }
  });
  
  if (backendCode) vm.runInContext(backendCode, ctx);
  return ctx;
}

function getMockSpreadsheetApp() {
  const { execFileSync } = require('child_process');
  
  function syncSupabase(table, method, body = null, query = '') {
    try {
      const SUPABASE_URL = process.env.SUPABASE_URL || 'https://kjcnotrxxthnzpgljeus.supabase.co';
      const SUPABASE_KEY = process.env.SUPABASE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtqY25vdHJ4eHRobnpwZ2xqZXVzIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MjY5ODY3MywiZXhwIjoyMDk4Mjc0NjczfQ.2yOpoM3C9ejhxA9hY0g88bkyU6KShhSaFHfnaBOLGiU';
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
        'Prefer': 'return=representation'
      };
      
      const args = [fullUrl, method, JSON.stringify(headers)];
      if (body) args.push(JSON.stringify(body));

      const out = execFileSync('node', ['-e', workerScript, ...args], { encoding: 'utf8', maxBuffer: 1024*1024*10 });
      const parts = out.split('|||');
      return parts[1] ? JSON.parse(parts[1]) : null;
    } catch (e) {
      console.error(`[SUPABASE ERROR ${table}]`, e.message);
      return [];
    }
  }

  const SHEET_HEADERS = {
    'Usuarios': ['ID','Email','Role','Name','Photo','Phone','Category','DNI','Birthdate','Age','JoinDate','BloodType','MedicalFit','ObraSocial','EmergencyContact','EmergencyPhone','ParentName','ParentPhone','Notes'],
    'Pagos': ['payment_id','email','month','amount','status','mp_link','collected_by','collected_at'],
    'Categorias': ['category_id','name','price_monthly','parent_category'],
    'Torneos': ['torneo_id','name','category','year','is_active'],
    'Finanzas_Torneos': ['movimiento_id','torneo_id','date','type','concept','amount','notes','created_by'],
    'Partidos': ['partido_id','torneo_id','date','opponent','is_home','status','goals_for','goals_against','notes'],
    'Admins': ['email','name','role']
  };

  const mockSpreadsheet = {
    getSheetByName: (name) => {
      const sheetName = name;
      let sbTable = sheetName.toLowerCase();
      if (sbTable === 'finanzas_torneos') sbTable = 'finanzas_torneos';
      
      return {
        getDataRange: () => ({
          getValues: () => {
            const sbRows = syncSupabase(sbTable, 'GET', null, '?select=*');
            if (!Array.isArray(sbRows)) return [];
            const headers = SHEET_HEADERS[sheetName] || Object.keys(sbRows[0] || {});
            const mappedRows = sbRows.map(obj => headers.map(h => obj[h.toLowerCase()] || ''));
            return [headers, ...mappedRows];
          }
        }),
        appendRow: (rowArray) => {
          const headers = SHEET_HEADERS[sheetName] || [];
          const obj = {};
          headers.forEach((h, idx) => {
            obj[h.toLowerCase()] = rowArray[idx] !== undefined ? rowArray[idx] : '';
          });
          syncSupabase(sbTable, 'POST', [obj]);
        },
        getLastRow: () => {
          const sbRows = syncSupabase(sbTable, 'GET', null, '?select=id');
          return Array.isArray(sbRows) ? sbRows.length + 1 : 1;
        },
        getRange: (row, col) => ({
          setValue: (val) => {
            // Simplified patch for fallback VM
            const headers = SHEET_HEADERS[sheetName] || [];
            const fieldName = (headers[col - 1] || '').toLowerCase();
            const sbRows = syncSupabase(sbTable, 'GET', null, '?select=*');
            if (!Array.isArray(sbRows)) return;
            const targetObj = sbRows[row - 2];
            if (targetObj) {
              const patchObj = {}; patchObj[fieldName] = val;
              let idCol = 'id';
              if (sheetName === 'Usuarios') idCol = 'email';
              else if (sheetName === 'Pagos') idCol = 'payment_id';
              syncSupabase(sbTable, 'PATCH', patchObj, `?${idCol}=eq.${encodeURIComponent(targetObj[idCol])}`);
            }
          },
          setNumberFormat: () => {}
        }),
        deleteRow: (rowIndex1Based) => {
          const sbRows = syncSupabase(sbTable, 'GET', null, '?select=*');
          if (!Array.isArray(sbRows)) return;
          const targetObj = sbRows[rowIndex1Based - 2];
          if (targetObj) {
            let idCol = 'id';
            if (sheetName === 'Usuarios') idCol = 'email';
            else if (sheetName === 'Pagos') idCol = 'payment_id';
            syncSupabase(sbTable, 'DELETE', null, `?${idCol}=eq.${encodeURIComponent(targetObj[idCol])}`);
          }
        }
      };
    }
  };
  return { openById: () => mockSpreadsheet };
}



app.post('/api/scan-receipt', async (req, res) => {
  try {
    const { imageBase64 } = req.body;
    if (!imageBase64) return res.status(400).json({ error: 'Falta imagen' });

    console.log('[OCR] Iniciando escaneo de comprobante...');
    const { data: { text } } = await Tesseract.recognize(imageBase64, 'spa', {
      logger: m => {} // silenciar logs internos
    });
    
    let code = null;
    const codeMatch = text.match(/(?:operaci[oó]n|comprobante|transacci[oó]n|cod.*?)\s*[:#Nro]*\s*([0-9A-Z]{8,20})/i);
    if (codeMatch) {
      code = codeMatch[1];
    }
    
    let amount = null;
    const amountMatch = text.match(/\$\s*([\d\.,]+)/);
    if (amountMatch) {
      amount = amountMatch[1];
    }
    
    console.log(`[OCR] Éxito. Codigo: ${code}, Monto: ${amount}`);
    res.json({ text, extractedCode: code, extractedAmount: amount });
  } catch (error) {
    console.error('[OCR Error]', error);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => console.log(`?? Servidor ERP profesional corriendo en puerto ${PORT}`));
