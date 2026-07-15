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
    if (false && apiFuncs[functionName]) {
      console.log(`[NATIVE] Ejecutando ${functionName} nativamente en Node.js`);
      const result = await apiFuncs[functionName](...(args || []));
      return res.json({ success: true, result });
    }
    
    // 2. Fallback a la version antigua de Codigo.js en la maquina virtual
    console.log(`[FALLBACK VM] Ejecutando ${functionName} en VM heredada`);
    const ctx = getSandboxContext();
    const func = ctx[functionName];
    if (typeof func !== 'function') {
      return res.status(404).json({ success: false, error: 'Funcion no encontrada' });
    }
    const result = func.apply(null, args || []);
    res.json({ success: true, result });
    
  } catch (err) {
    console.error(`[API RUN ERROR] ${functionName}:`, err);
    res.status(500).json({ success: false, error: 'Error interno del servidor' });
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
    'Usuarios': ['ID','Username','Role','Name','Photo','Phone','Category','DNI','Birthdate','Age','JoinDate','BloodType','MedicalFit','ObraSocial','EmergencyContact','EmergencyPhone','ParentName','ParentPhone','Notes'],
    'Pagos': ['payment_id','username','month','amount','status','mp_link','collected_by','collected_at'],
    'Categorias': ['category_id','name','price_monthly','parent_category'],
    'Torneos': ['torneo_id','name','category','year','is_active'],
    'Finanzas_Torneos': ['movimiento_id','torneo_id','date','type','concept','amount','notes','created_by'],
    'Partidos': ['partido_id','torneo_id','date','opponent','is_home','status','goals_for','goals_against','notes'],
    'Admins': ['username','name','role']
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
            let colName = h.toLowerCase();
            if (sbTable === 'pagos' && colName === 'username') colName = 'email';
            if (sbTable === 'logs_audit' && colName === 'username') colName = 'user_email';
            obj[colName] = rowArray[idx] !== undefined ? rowArray[idx] : '';
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
            let fieldName = (headers[col - 1] || '').toLowerCase();
            const sbRows = syncSupabase(sbTable, 'GET', null, '?select=*');
            if (!Array.isArray(sbRows)) return;
            const targetObj = sbRows[row - 2];
            if (targetObj) {
              if (sbTable === 'pagos' && fieldName === 'username') fieldName = 'email';
              if (sbTable === 'logs_audit' && fieldName === 'username') fieldName = 'user_email';
              const patchObj = {}; patchObj[fieldName] = val;
              let idCol = 'id';
              if (sheetName === 'Usuarios') idCol = 'username';
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
            if (sheetName === 'Usuarios') idCol = 'username';
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
    const codeMatch = text.match(/(?:operaci[o�]n|comprobante|transacci[o�]n|cod.*?)\s*[:#Nro]*\s*([0-9A-Z]{8,20})/i);
    if (codeMatch) {
      code = codeMatch[1];
    }
    
    let amount = null;
    const amountMatch = text.match(/\$\s*([\d\.,]+)/);
    if (amountMatch) {
      amount = amountMatch[1];
    }
    
    console.log(`[OCR] �xito. Codigo: ${code}, Monto: ${amount}`);
    res.json({ text, extractedCode: code, extractedAmount: amount });
  } catch (error) {
    console.error('[OCR Error]', error);
    res.status(500).json({ error: error.message });
  }
});
// --- NOTIFICACIONES PUSH E IN-APP ---
const webpush = require('web-push');

// Configuración de VAPID keys para Web Push (se usarán variables de entorno en producción)
const vapidPublicKey = process.env.VAPID_PUBLIC_KEY || 'BFvdjHU0w3bplc8vT-ywJwYGy_hQmJqUNxvXecR9GxMA97-ItmDLlxNDYkUe88CdFPfMdDat6Tl535-DVel5vSc';
const vapidPrivateKey = process.env.VAPID_PRIVATE_KEY || 'uA05adM20cxNTUFDyo33zZCcpc0gf2ZJK17e_QojYP0';
webpush.setVapidDetails(
  'mailto:admin@haedofutsal.com',
  vapidPublicKey,
  vapidPrivateKey
);

// 1. Guardar suscripción del dispositivo
app.post('/api/notifications/subscribe', auth.authenticateToken, async (req, res) => {
  const subscription = req.body;
  const username_socio = req.user.username;
  
  if (!username_socio) return res.status(401).json({ error: 'No autorizado' });

  // Guardar en Supabase (upsert)
  try {
    const { error } = await supabase
      .from('suscripciones_push')
      .upsert({ username_socio: username_socio, suscripcion: subscription }, { onConflict: 'username_socio' });
      
    if (error) throw error;
    res.status(201).json({ success: true });
  } catch (err) {
    console.error('Error al suscribir:', err);
    res.status(500).json({ error: 'Error al guardar suscripción' });
  }
});

// 2. Enviar Notificación (Solo Admin)
app.post('/api/notifications/send', auth.authenticateToken, async (req, res) => {
  if (req.user.role !== 'Admin' && req.user.role !== 'Entrenador' && req.user.role !== 'Tesorero') {
    return res.status(403).json({ error: 'Permisos insuficientes' });
  }

  const { titulo, mensaje, categoria_destino, estado_pago_destino } = req.body;

  try {
    // A) Guardar en BD para las Alertas In-App
    const { data: nuevaNotif, error: errBd } = await supabase
      .from('notificaciones')
      .insert([{
        titulo,
        mensaje,
        categoria_destino: categoria_destino || 'TODAS',
        estado_pago_destino: estado_pago_destino || 'TODOS'
      }]);
      
    if (errBd) throw errBd;

    // B) Disparar Push a los dispositivos relevantes
    const { data: subs, error: errSub } = await supabase.from('suscripciones_push').select('*');
    
    if (!errSub && subs) {
      const payload = JSON.stringify({ title: titulo, body: mensaje });
      const promises = subs.map(sub => {
        return webpush.sendNotification(sub.suscripcion, payload).catch(err => {
          console.error("Fallo al enviar push a ", sub.username_socio, err);
        });
      });
      await Promise.all(promises);
    }
    
    res.status(200).json({ success: true, message: 'Notificaciones enviadas con éxito' });
  } catch (err) {
    console.error('Error al enviar notificaciones:', err);
    res.status(500).json({ error: 'Error interno del servidor' });
  }
});

// === MERCADO PAGO ENDPOINT ===
app.get('/api/mp-transfers', async (req, res) => {
  try {
    let MP_TOKEN = process.env.MP_ACCESS_TOKEN;
    if (!MP_TOKEN || MP_TOKEN === "MOCK_MP_ACCESS_TOKEN_DEVELOPMENT") {
        MP_TOKEN = "APP_USR-3322444120483456-062819-f186f817a6a28fd7251c13baaf3d014e-43153257";
    }
    
    let queryParams = new URLSearchParams();
    queryParams.append('sort', 'date_created');
    queryParams.append('criteria', 'desc');
    queryParams.append('limit', '1000');

    // add date filters if requested
    const startDate = req.query.startDate;
    const endDate = req.query.endDate;
    if (startDate && endDate) {
        queryParams.append('range', 'date_created');
        queryParams.append('begin_date', new Date(startDate).toISOString());
        queryParams.append('end_date', new Date(endDate + 'T23:59:59.999Z').toISOString());
    }

    const axios = require('axios');
    const mpRes = await axios.get('https://api.mercadopago.com/v1/payments/search?' + queryParams.toString(), {
       headers: {
         'Authorization': 'Bearer ' + MP_TOKEN
       }
    });

    res.json(mpRes.data);
  } catch (error) {
    console.error('Error fetching MP transfers:', error.response ? error.response.data : error.message);
    res.status(500).json({ error: error.response ? JSON.stringify(error.response.data) : error.message });
  }
});


// === RUTAS DE CONCILIACIÓN DE PAGOS ===

app.post('/api/payments/reconcile', async (req, res) => {
  try {
    const { paymentId, username, amount, month, transferMethod, base64Receipt, transactionId, ocrAmount, ocrText } = req.body;
    let MP_TOKEN = process.env.MP_ACCESS_TOKEN;
    if (!MP_TOKEN || MP_TOKEN === "MOCK_MP_ACCESS_TOKEN_DEVELOPMENT") {
        MP_TOKEN = "APP_USR-3322444120483456-062819-f186f817a6a28fd7251c13baaf3d014e-43153257";
    }

    // 1. Get payment from Supabase
    const { data: pagos, error: errP } = await supabase.from('pagos').select('*').eq('payment_id', paymentId);
    if (errP || !pagos || pagos.length === 0) {
      return res.status(404).json({ success: false, message: "No se encontró el registro de pago." });
    }
    const pago = pagos[0];

    // Clean amounts
    const targetAmount = parseFloat(amount.toString().replace(/[^0-9.-]+/g,""));
    
    let cleanTxId = (transactionId || "").trim();
    const isCoelsaId = cleanTxId && (cleanTxId.length === 22 || /[a-zA-Z]/.test(cleanTxId));

    let casoAExitoso = false;
    let fallbackToCasoB = false;

    if (cleanTxId && !isCoelsaId) {
      // Check if already used
      const { data: usedPagos } = await supabase.from('pagos').select('payment_id').ilike('collected_by', `%${cleanTxId}%`);
      if (usedPagos && usedPagos.length > 0) {
         return res.json({ success: false, message: `El comprobante ingresado (ID: ${cleanTxId}) ya fue utilizado para acreditar otra cuota.` });
      }

      // Check MP
      const axios = require('axios');
      try {
        const mpRes = await axios.get(`https://api.mercadopago.com/v1/payments/${cleanTxId}`, {
          headers: { 'Authorization': 'Bearer ' + MP_TOKEN }
        });
        if (mpRes.data && mpRes.data.status === 'approved') {
          const mpAmount = parseFloat(mpRes.data.transaction_amount);
          if (Math.abs(mpAmount - targetAmount) < 10.0) {
            casoAExitoso = true;
          } else {
            fallbackToCasoB = true;
          }
        } else {
          fallbackToCasoB = true;
        }
      } catch (e) {
        fallbackToCasoB = true;
      }
    }

    if (!casoAExitoso && (!cleanTxId || isCoelsaId || fallbackToCasoB)) {
      if (cleanTxId) {
        const { data: usedPagos } = await supabase.from('pagos').select('payment_id').ilike('collected_by', `%${cleanTxId}%`);
        if (usedPagos && usedPagos.length > 0) {
           return res.json({ success: false, message: `El comprobante ingresado (ID/Coelsa: ${cleanTxId}) ya fue utilizado para acreditar otra cuota.` });
        }
      }

      // Caso B: manual review required
      let collectedByStr = transferMethod;
      if (cleanTxId) collectedByStr += " ID:" + cleanTxId;
      
      const updateData = {
        status: 'En Revisión',
        mp_link: base64Receipt || "",
        collected_by: collectedByStr,
        collected_at: new Date().toISOString()
      };
      
      await supabase.from('pagos').update(updateData).eq('payment_id', paymentId);
      return res.json({ success: false, casoB: true, message: "Validación manual requerida" });
    }

    // Caso A: Success
    const updateData = {
      status: 'Pagado',
      mp_link: '',
      collected_by: `MercadoPago (ID:${cleanTxId})`,
      collected_at: new Date().toISOString()
    };
    await supabase.from('pagos').update(updateData).eq('payment_id', paymentId);

    return res.json({ success: true, message: "El comprobante fue procesado e imputado exitosamente a tu cuenta de forma automática." });

  } catch (error) {
    console.error('Error in reconcile:', error);
    res.status(500).json({ error: 'Error interno en la conciliación' });
  }
});

app.post('/api/payments/request-review', async (req, res) => {
  try {
    const { paymentId, username, amount, month, paymentMethod, base64Receipt, failReason } = req.body;
    const updateData = {
      status: 'En Revisión',
      mp_link: base64Receipt || "",
      collected_by: paymentMethod + (failReason ? " (" + failReason + ")" : ""),
      collected_at: new Date().toISOString()
    };
    await supabase.from('pagos').update(updateData).eq('payment_id', paymentId);
    return res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/payments/approve', async (req, res) => {
  try {
    const { paymentId, username, amount, month } = req.body;
    const updateData = {
      status: 'Pagado',
      mp_link: '',
      collected_at: new Date().toISOString()
    };
    // Prepend 'Aprobado Manual - ' to collected_by
    const { data: pagos } = await supabase.from('pagos').select('collected_by').eq('payment_id', paymentId);
    if (pagos && pagos.length > 0) {
       updateData.collected_by = "Aprobado Manual - " + pagos[0].collected_by;
    }
    await supabase.from('pagos').update(updateData).eq('payment_id', paymentId);
    return res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/payments/reject', async (req, res) => {
  try {
    const { paymentId, username, month, motivo } = req.body;
    const updateData = {
      status: 'Deuda',
      mp_link: '',
      collected_by: '',
      collected_at: null
    };
    await supabase.from('pagos').update(updateData).eq('payment_id', paymentId);
    
    // Podría insertarse en alguna tabla de auditoría, pero por ahora lo dejamos simple.
    return res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/ping', (req, res) => res.status(200).send('Pong! Servidor activo.'));

app.listen(PORT, () => {
  console.log('Servidor ERP profesional corriendo en puerto ' + PORT);
  const RENDER_URL = process.env.RENDER_EXTERNAL_URL;
  if (RENDER_URL) {
    const https = require('https');
    setInterval(() => {
      https.get(RENDER_URL + '/ping', (resp) => {
        if (resp.statusCode === 200) console.log('Keep-Alive ping exitoso.');
      }).on('error', (err) => console.error('Keep-Alive ping fallido:', err.message));
    }, 840000);
    console.log('Keep-Alive activado hacia: ' + RENDER_URL + '/ping');
  }
});
