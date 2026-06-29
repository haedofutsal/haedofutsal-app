/**
 * Google Apps Script - Backend API para ERP Web Futsal Haedo (Ampliando Torneos, Finanzas y Socios)
 * Desarrollado con sintaxis moderna ES6 (motor V8).
 * 
 * Este archivo actúa como el backend que sirve el frontend HTML (HtmlService)
 * y expone las funciones del servidor llamadas de forma segura desde el cliente
 * mediante "google.script.run".
 */

// ==========================================
// CONFIGURACIÓN Y CONSTANTES
// ==========================================
const SPREADSHEET_ID = "1nDZWIxdGPWK8YY86JcOW9YCwTuApQX380H7Wk4JGjWg"; // Planilla oficial de Futsal Haedo
const HOJA_PAGOS = "Pagos";
const HOJA_USUARIOS = "Usuarios";
const HOJA_CATEGORIAS = "Categorias";
const HOJA_ADMINS = "Admins";
const HOJA_TORNEOS = "Torneos";
const HOJA_FINANZAS = "Finanzas_Torneos";
const HOJA_PARTIDOS = "Partidos";
const HOJA_LOGS = "Logs_Audit";

/**
 * Endpoint POST para sincronización de API externa (Node.js / Render)
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const functionName = postData.functionName;
    const args = postData.args || [];
    
    if (typeof this[functionName] === 'function') {
      const result = this[functionName].apply(this, args);
      return ContentService.createTextOutput(JSON.stringify({ success: true, result: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Función no encontrada: " + functionName }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Obtiene el token de acceso de Mercado Pago desde las propiedades del script.
 */
function getMercadoPagoToken() {
  const token = PropertiesService.getScriptProperties().getProperty("MP_ACCESS_TOKEN");
  if (!token) {
    console.error("ERROR: No se encontró la propiedad 'MP_ACCESS_TOKEN' en PropertiesService.");
    throw new Error("Credenciales de Mercado Pago no configuradas.");
  }
  return token;
}

/**
 * Sirve la interfaz de usuario HTML (Index.html) cuando se ingresa a la Web App.
 */
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile("Index");
    return template.evaluate()
      .setTitle("Futsal Haedo - ERP Deportivo")
      .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error("Error sirviendo la página web:", error);
    return HtmlService.createHtmlOutput(`<h3>Error cargando el sistema: ${error.message}</h3>`);
  }
}

/**
 * Permite incluir sub-archivos si se divide la interfaz.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// MÉTODOS DE LA API (Expuestos a google.script.run)
// ==========================================

/**
 * Realiza el inicio de sesión verificando si el email existe en la base de datos de Socios o Admins.
 */
function loginUsuario(email) {
  try {
    if (!email) throw new Error("El correo es requerido.");
    const emailNorm = email.toLowerCase().trim();
    const ss = getSpreadsheet();
    
    // 1. Verificar si es Administrador en la tabla "Admins"
    const sheetAdmins = ss.getSheetByName(HOJA_ADMINS);
    if (sheetAdmins) {
      const adminsList = sheetAdmins.getDataRange().getValues().slice(1).map(row => row[0].toString().toLowerCase().trim());
      if (adminsList.includes(emailNorm)) {
        return {
          success: true,
          role: "Admin",
          email: emailNorm,
          name: "Administrador Futsal Haedo"
        };
      }
    }
    
    // 2. Verificar si es Socio (Deportista) en la tabla "Usuarios"
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheetUsers) throw new Error("La tabla de usuarios no está inicializada.");
    
    const usersData = sheetUsers.getDataRange().getValues();
    const headers = usersData[0];
    const emailColIndex = headers.indexOf("Email");
    
    if (emailColIndex === -1) throw new Error("No se encontró la columna 'Email' en la hoja de Usuarios.");
    
    let userRow = null;
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][emailColIndex].toString().toLowerCase().trim() === emailNorm) {
        userRow = usersData[i];
        break;
      }
    }
    
    if (!userRow) {
      return { success: false, message: "El correo electrónico no se encuentra registrado en el sistema." };
    }
    
    const userData = {};
    headers.forEach((header, index) => {
      userData[header] = userRow[index];
    });
    
    return {
      success: true,
      role: userData.Role || "Deportista",
      email: emailNorm,
      name: userData.Name || "Socio",
      user: userData
    };
    
  } catch (error) {
    console.error("Error en loginUsuario:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Obtiene la información personal, el historial de pagos y partidos de la categoría de un deportista.
 */
function obtenerDatosSocio(email) {
  try {
    const emailNorm = email.toLowerCase().trim();
    const ss = getSpreadsheet();
    const getVal = (row, idx) => (row && idx !== -1 && idx < row.length && row[idx] !== undefined && row[idx] !== null) ? row[idx].toString().trim() : "";
    
    // 1. Obtener datos del perfil del deportista
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    const usersData = sheetUsers.getDataRange().getValues();
    const userHeaders = usersData[0];
    const emailColIndex = userHeaders.indexOf("Email");
    
    let perfil = {};
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      if (getVal(row, emailColIndex).toLowerCase() === emailNorm) {
        userHeaders.forEach((header, index) => {
          perfil[header] = getVal(row, index);
        });
        break;
      }
    }
    
    const socioCategory = perfil.Category || "";
    
    // 2. Obtener historial de pagos
    const sheetPagos = ss.getSheetByName(HOJA_PAGOS);
    const pagosData = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    const pagosHeaders = pagosData[0] || [];
    let pagosEmailColIndex = pagosHeaders.indexOf("Email");
    if (pagosEmailColIndex === -1) pagosEmailColIndex = pagosHeaders.indexOf("User_Email");
    
    const misPagos = [];
    if (pagosEmailColIndex !== -1) {
      for (let i = 1; i < pagosData.length; i++) {
        const row = pagosData[i];
        if (!getVal(row, pagosHeaders.indexOf("Payment_ID"))) continue;
        if (getVal(row, pagosEmailColIndex).toLowerCase() === emailNorm) {
          const pago = {};
          pagosHeaders.forEach((header, index) => {
            pago[header] = getVal(row, index);
          });
          misPagos.push(pago);
        }
      }
    }
    misPagos.sort((a, b) => (b.Month || "").localeCompare(a.Month || ""));
    
    // 3. Obtener partidos programados para la categoría del deportista
    const misPartidos = [];
    const sheetTorneos = ss.getSheetByName(HOJA_TORNEOS);
    const sheetPartidos = ss.getSheetByName(HOJA_PARTIDOS);
    
    if (sheetTorneos && sheetPartidos && socioCategory !== "") {
      const torneosData = sheetTorneos.getDataRange().getValues();
      const tHeaders = torneosData[0];
      const tIdCol = tHeaders.indexOf("Torneo_ID");
      const tCatCol = tHeaders.indexOf("Category");
      const tNameCol = tHeaders.indexOf("Name");
      
      // Obtener torneos que disputa su categoría
      const misTorneosIds = [];
      const torneosMap = {};
      for (let i = 1; i < torneosData.length; i++) {
        if (getVal(torneosData[i], tCatCol) === socioCategory) {
          const tId = getVal(torneosData[i], tIdCol);
          misTorneosIds.push(tId);
          torneosMap[tId] = getVal(torneosData[i], tNameCol);
        }
      }
      
      // Cargar partidos de esos torneos
      const partidosData = sheetPartidos.getDataRange().getValues();
      const pHeaders = partidosData[0];
      const pIdCol = pHeaders.indexOf("Partido_ID");
      const pTIdCol = pHeaders.indexOf("Torneo_ID");
      const pDateCol = pHeaders.indexOf("Date");
      const pOppCol = pHeaders.indexOf("Opponent");
      const pLocCol = pHeaders.indexOf("Location");
      const pResCol = pHeaders.indexOf("Result");
      
      for (let i = 1; i < partidosData.length; i++) {
        const row = partidosData[i];
        if (!getVal(row, pIdCol)) continue;
        const torneoId = getVal(row, pTIdCol);
        
        if (misTorneosIds.includes(torneoId)) {
          const partidoObj = {};
          pHeaders.forEach((h, idx) => {
            partidoObj[h] = getVal(row, idx);
          });
          partidoObj.Torneo_Name = torneosMap[torneoId] || "Torneo";
          misPartidos.push(partidoObj);
        }
      }
    }
    
    // Ordenar partidos por fecha
    misPartidos.sort((a, b) => (b.Date || "").localeCompare(a.Date || ""));
    
    return {
      success: true,
      perfil,
      pagos: misPagos,
      partidos: misPartidos
    };
    
  } catch (error) {
    console.error("Error en obtenerDatosSocio:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Retorna métricas generales financieras, deudores, socios, partidos y balances de torneos.
 */
function obtenerDatosAdmin() {
  try {
    const ss = getSpreadsheet();
    const getVal = (row, idx) => (row && idx !== -1 && idx < row.length && row[idx] !== undefined && row[idx] !== null) ? row[idx].toString().trim() : "";
    
    // 1. Cargar Usuarios
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    const usersData = sheetUsers ? sheetUsers.getDataRange().getValues() : [];
    const userHeaders = usersData[0] || [];
    const rowsUsers = usersData.slice(1);
    
    const userEmailColIndex = userHeaders.indexOf("Email");
    const roleColIndex = userHeaders.indexOf("Role");
    const nameColIndex = userHeaders.indexOf("Name");
    const phoneColIndex = userHeaders.indexOf("Phone");
    const catColIndex = userHeaders.indexOf("Category");
    
    const listaSocios = [];
    rowsUsers.forEach(row => {
      const email = getVal(row, userEmailColIndex);
      const role = getVal(row, roleColIndex);
      if (email === "" || role !== "Deportista") return;
      
      const socioObj = {};
      userHeaders.forEach((h, idx) => {
        socioObj[h] = getVal(row, idx);
      });
      if (!socioObj.Photo) {
        socioObj.Photo = "https://i.pravatar.cc/150?u=" + email.toLowerCase().replace(/[^a-zA-Z0-9]/g, "");
      }
      listaSocios.push(socioObj);
    });
    
    // 2. Cargar Pagos
    const sheetPagos = ss.getSheetByName(HOJA_PAGOS);
    const pagosData = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    const pagosHeaders = pagosData[0] || [];
    const rowsPagos = pagosData.slice(1);
    
    const idColIndex = pagosHeaders.indexOf("Payment_ID");
    const statusColIndex = pagosHeaders.indexOf("Status");
    const amountColIndex = pagosHeaders.indexOf("Amount");
    const monthColIndex = pagosHeaders.indexOf("Month");
    let emailColIndex = pagosHeaders.indexOf("Email");
    if (emailColIndex === -1) emailColIndex = pagosHeaders.indexOf("User_Email");
    const linkColIndex = pagosHeaders.indexOf("MP_Link");
    const collectedByColIndex = pagosHeaders.indexOf("Collected_By");
    const collectedAtColIndex = pagosHeaders.indexOf("Collected_At");
    
    let recaudacionTotal = 0;
    let deudaPendiente = 0;
    const deudoresMap = {};
    const todosLosPagos = [];
    
    rowsPagos.forEach(row => {
      const id = getVal(row, idColIndex);
      if (!id) return;
      
      const email = getVal(row, emailColIndex).toLowerCase();
      const status = getVal(row, statusColIndex);
      const amount = parseFloat(getVal(row, amountColIndex)) || 0;
      const month = getVal(row, monthColIndex);
      
      const socio = listaSocios.find(s => s.Email.toLowerCase() === email);
      const socioName = socio ? socio.Name : "Socio Desconocido";
      
      const pagoObj = {
        Payment_ID: id,
        Email: email,
        SocioName: socioName,
        Month: month,
        Amount: amount,
        Status: status,
        MP_Link: getVal(row, linkColIndex),
        Collected_By: getVal(row, collectedByColIndex),
        Collected_At: getVal(row, collectedAtColIndex)
      };
      todosLosPagos.push(pagoObj);
      
      if (status === "Pagado") {
        recaudacionTotal += amount;
      } else if (status === "Pendiente") {
        deudaPendiente += amount;
        if (!deudoresMap[email]) {
          deudoresMap[email] = {
            Name: socioName,
            Email: email,
            CuotasPendientes: [],
            TotalDeuda: 0
          };
        }
        deudoresMap[email].CuotasPendientes.push(month);
        deudoresMap[email].TotalDeuda += amount;
      }
    });
    
    const listaDeudores = Object.values(deudoresMap).sort((a, b) => b.TotalDeuda - a.TotalDeuda);
    todosLosPagos.sort((a, b) => (b.Month || "").localeCompare(a.Month || ""));
    
    // 3. Cargar Torneos y sus Finanzas
    const sheetTorneos = ss.getSheetByName(HOJA_TORNEOS);
    const sheetFinanzas = ss.getSheetByName(HOJA_FINANZAS);
    
    const torneos = [];
    const torneosData = sheetTorneos ? sheetTorneos.getDataRange().getValues() : [];
    const tHeaders = torneosData[0] || [];
    const rowsTorneos = torneosData.slice(1);
    
    const tIdCol = tHeaders.indexOf("Torneo_ID");
    const tNameCol = tHeaders.indexOf("Name");
    const tCatCol = tHeaders.indexOf("Category");
    
    rowsTorneos.forEach(row => {
      const tId = getVal(row, tIdCol);
      if (!tId) return;
      torneos.push({
        Torneo_ID: tId,
        Name: getVal(row, tNameCol),
        Category: getVal(row, tCatCol),
        Ingresos: 0,
        Gastos: 0,
        Balance: 0
      });
    });
    
    const finanzasData = sheetFinanzas ? sheetFinanzas.getDataRange().getValues() : [];
    const fHeaders = finanzasData[0] || [];
    const rowsFinanzas = finanzasData.slice(1);
    
    const fIdCol = fHeaders.indexOf("Movimiento_ID");
    const fTIdCol = fHeaders.indexOf("Torneo_ID");
    const fTypeCol = fHeaders.indexOf("Type");
    const fConceptCol = fHeaders.indexOf("Concept");
    const fAmountCol = fHeaders.indexOf("Amount");
    const fDateCol = fHeaders.indexOf("Date");
    
    const todosLosMovimientos = [];
    
    rowsFinanzas.forEach(row => {
      const fId = getVal(row, fIdCol);
      if (!fId) return;
      
      const tId = getVal(row, fTIdCol);
      const type = getVal(row, fTypeCol);
      const amount = parseFloat(getVal(row, fAmountCol)) || 0;
      const concept = getVal(row, fConceptCol);
      const date = getVal(row, fDateCol);
      
      const torneo = torneos.find(t => t.Torneo_ID === tId);
      const torneoName = torneo ? torneo.Name : "General";
      
      todosLosMovimientos.push({
        Movimiento_ID: fId,
        Torneo_ID: tId,
        Torneo_Name: torneoName,
        Type: type,
        Concept: concept,
        Amount: amount,
        Date: date
      });
      
      if (torneo) {
        if (type === "Ingreso") {
          torneo.Ingresos += amount;
        } else if (type === "Gasto") {
          torneo.Gastos += amount;
        }
        torneo.Balance = torneo.Ingresos - torneo.Gastos;
      }
    });
    todosLosMovimientos.sort((a, b) => (b.Date || "").localeCompare(a.Date || ""));
    
    // 4. Cargar Partidos
    const sheetPartidos = ss.getSheetByName(HOJA_PARTIDOS);
    const partidosData = sheetPartidos ? sheetPartidos.getDataRange().getValues() : [];
    const pHeaders = partidosData[0] || [];
    const rowsPartidos = partidosData.slice(1);
    
    const pIdCol = pHeaders.indexOf("Partido_ID");
    const pTIdCol = pHeaders.indexOf("Torneo_ID");
    const pDateCol = pHeaders.indexOf("Date");
    const pOppCol = pHeaders.indexOf("Opponent");
    const pLocCol = pHeaders.indexOf("Location");
    const pResCol = pHeaders.indexOf("Result");
    
    const todosLosPartidos = [];
    rowsPartidos.forEach(row => {
      const pId = getVal(row, pIdCol);
      if (!pId) return;
      
      const tId = getVal(row, pTIdCol);
      const torneo = torneos.find(t => t.Torneo_ID === tId);
      
      const partidoObj = {};
      pHeaders.forEach((h, idx) => {
        partidoObj[h] = getVal(row, idx);
      });
      partidoObj.Torneo_Name = torneo ? torneo.Name : "Desconocido";
      todosLosPartidos.push(partidoObj);
    });
    todosLosPartidos.sort((a, b) => (b.Date || "").localeCompare(a.Date || ""));
    
    return {
      success: true,
      metrics: {
        totalSocios: listaSocios.length,
        recaudacionTotal,
        deudaPendiente
      },
      socios: listaSocios,
      deudores: listaDeudores,
      pagos: todosLosPagos,
      torneos: torneos,
      finanzas: todosLosMovimientos,
      partidos: todosLosPartidos
    };
    
  } catch (error) {
    console.error("Error en obtenerDatosAdmin:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Permite al administrador consultar los pagos históricos de un socio particular.
 */
function obtenerFichaSocioCompleta(email) {
  try {
    const ss = getSpreadsheet();
    const getVal = (row, idx) => (row && idx !== -1 && idx < row.length && row[idx] !== undefined && row[idx] !== null) ? row[idx].toString().trim() : "";
    
    // 1. Perfil
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    const usersData = sheetUsers.getDataRange().getValues();
    const userHeaders = usersData[0];
    const emailColIndex = userHeaders.indexOf("Email");
    
    let perfil = {};
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      if (getVal(row, emailColIndex).toLowerCase() === email.toLowerCase().trim()) {
        userHeaders.forEach((header, index) => {
          perfil[header] = getVal(row, index);
        });
        break;
      }
    }
    
    // 2. Historial de Pagos completo
    const sheetPagos = ss.getSheetByName(HOJA_PAGOS);
    const pagosData = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    const pagosHeaders = pagosData[0] || [];
    let pagosEmailColIndex = pagosHeaders.indexOf("Email");
    if (pagosEmailColIndex === -1) pagosEmailColIndex = pagosHeaders.indexOf("User_Email");
    
    const historialPagos = [];
    if (pagosEmailColIndex !== -1) {
      for (let i = 1; i < pagosData.length; i++) {
        const row = pagosData[i];
        if (!getVal(row, pagosHeaders.indexOf("Payment_ID"))) continue;
        if (getVal(row, pagosEmailColIndex).toLowerCase() === email.toLowerCase().trim()) {
          const pago = {};
          pagosHeaders.forEach((header, index) => {
            pago[header] = getVal(row, index);
          });
          historialPagos.push(pago);
        }
      }
    }
    historialPagos.sort((a, b) => (b.Month || "").localeCompare(a.Month || ""));
    
    return {
      success: true,
      perfil,
      pagos: historialPagos
    };
  } catch (error) {
    console.error("Error en obtenerFichaSocioCompleta:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Registra un nuevo deportista / socio en la base de datos (Hoja Usuarios).
 */
function registrarSocioNuevo(socioObj, userEmail) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheet) throw new Error("No se encontró la hoja de usuarios.");

    const email = (socioObj.Email || "").trim().toLowerCase();
    if (!email) throw new Error("El correo electrónico es obligatorio.");

    // Verificar si ya existe
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailColIdx = headers.indexOf("Email");

    for (let i = 1; i < data.length; i++) {
      if ((data[i][emailColIdx] || "").toString().toLowerCase().trim() === email) {
        throw new Error("Ya existe un deportista registrado con el correo " + email);
      }
    }

    // Calcular edad si se proporcionó fecha de nacimiento
    let age = "";
    if (socioObj.BirthDate) {
      const birth = new Date(socioObj.BirthDate);
      const today = new Date();
      age = today.getFullYear() - birth.getFullYear();
      const m = today.getMonth() - birth.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < birth.getDate())) {
        age--;
      }
    }

    const todayStr = new Date().toISOString().split("T")[0];
    const isMinor = (socioObj.Category || "").includes("Bab") || (socioObj.Category || "").includes("Juv") || (parseInt(age || "20") < 18);
    const assignedRole = socioObj.Role || "Deportista";

    const rowArray = [
      email,
      assignedRole,
      socioObj.Name || "",
      socioObj.Phone || "",
      socioObj.Category || "Fut-May+35",
      socioObj.DNI || "",
      socioObj.BirthDate || "",
      age ? age.toString() : "",
      todayStr,
      socioObj.BloodType || "O+",
      socioObj.MedicalFit || "Apto Físico Vigente",
      socioObj.ObraSocial || "Particular",
      socioObj.EmergencyContact || socioObj.ParentName || "",
      socioObj.EmergencyPhone || socioObj.ParentPhone || "",
      isMinor ? (socioObj.ParentName || "") : "",
      isMinor ? (socioObj.ParentPhone || "") : "",
      socioObj.Notes || "",
      "" // Photo URL
    ];

    sheet.appendRow(rowArray);

    // Si es Administrador, asegurar que también figure en la hoja Admins
    if (assignedRole === "Admin") {
      const sheetAdmins = ss.getSheetByName(HOJA_ADMINS);
      if (sheetAdmins) {
        const adminsData = sheetAdmins.getDataRange().getValues();
        const existing = adminsData.map(r => (r[0] || "").toString().toLowerCase().trim());
        if (!existing.includes(email)) {
          sheetAdmins.appendRow([email, socioObj.Name || "Admin"]);
        }
      }
    }

    if (typeof registrarLogAuditoria === "function") {
      registrarLogAuditoria(userEmail || "admin@futsalhaedo.com", "CREAR", "PADRON_SOCIOS", `Dado de alta usuario: ${socioObj.Name || email} (${email}) con rol ${assignedRole}`);
    }

    return { success: true, message: "Socio registrado con éxito." };
  } catch (err) {
    console.error("Error en registrarSocioNuevo:", err);
    return { success: false, message: err.message };
  }
}

/**
 * Registra una entrada de auditoría en el Log General de la App.
 */
function registrarLogAuditoria(userEmail, actionType, entity, details) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(HOJA_LOGS);
    if (!sheet) {
      sheet = ss.insertSheet(HOJA_LOGS);
      sheet.appendRow(["Log_ID", "Timestamp", "User_Email", "Action_Type", "Entity", "Details"]);
    }
    const logId = `LOG-${Date.now().toString().slice(-6)}`;
    const timestamp = new Date().toISOString().replace("T", " ").substring(0, 19);
    sheet.appendRow([logId, timestamp, userEmail || "Admin", actionType, entity, details]);
    SpreadsheetApp.flush();
  } catch (err) {
    console.error("Error al registrar log de auditoría:", err);
  }
}

/**
 * Registra un movimiento financiero asociado a un Torneo (Ingreso/Gasto/Sponsor).
 */
function registrarMovimientoTorneo(movObj, userEmail) {
  try {
    const { Torneo_ID, Type, Concept, Amount, Payment_Method } = movObj;
    if (!Torneo_ID || !Type || !Concept || !Amount) {
      throw new Error("Faltan parámetros obligatorios para registrar el movimiento.");
    }
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_FINANZAS);
    if (!sheet) throw new Error("No se encontró la hoja de finanzas.");
    
    const movId = `MOV-${Date.now().toString().slice(-6)}`;
    const dateFormatted = new Date().toISOString().split("T")[0]; // YYYY-MM-DD
    const method = Payment_Method || "Efectivo";
    
    // Columnas: Movimiento_ID, Torneo_ID, Type, Concept, Amount, Date, Payment_Method
    sheet.appendRow([movId, Torneo_ID, Type, Concept, parseFloat(Amount), dateFormatted, method]);
    
    const lastRowIndex = sheet.getLastRow();
    if (sheet.getRange(lastRowIndex, 5).setNumberFormat) {
      sheet.getRange(lastRowIndex, 5).setNumberFormat("$#,##0.00");
    }
    SpreadsheetApp.flush();
    
    registrarLogAuditoria(userEmail || "Admin", "CREAR", "MOVIMIENTO", `Registrado movimiento ID ${movId}: "${Concept}" por $${parseFloat(Amount).toLocaleString('es-AR')} (${method})`);
    
    return { success: true, message: "Movimiento financiero registrado correctamente." };
    
  } catch (error) {
    console.error("Error en registrarMovimientoTorneo:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Edita un movimiento financiero existente.
 */
function editarMovimientoTorneo(movObj, userEmail) {
  try {
    const { Movimiento_ID, Torneo_ID, Type, Concept, Amount, DateStr, Payment_Method } = movObj;
    if (!Movimiento_ID || !Concept || !Amount) {
      throw new Error("Faltan parámetros obligatorios para editar el movimiento.");
    }
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_FINANZAS);
    if (!sheet) throw new Error("No se encontró la hoja de finanzas.");
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const movIdCol = headers.indexOf("Movimiento_ID");
    
    let updated = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][movIdCol].toString() === Movimiento_ID.toString()) {
        const rowNum = i + 1;
        if (Torneo_ID) sheet.getRange(rowNum, headers.indexOf("Torneo_ID") + 1).setValue(Torneo_ID);
        if (Type) sheet.getRange(rowNum, headers.indexOf("Type") + 1).setValue(Type);
        if (Concept) sheet.getRange(rowNum, headers.indexOf("Concept") + 1).setValue(Concept);
        if (Amount) sheet.getRange(rowNum, headers.indexOf("Amount") + 1).setValue(parseFloat(Amount));
        if (DateStr) sheet.getRange(rowNum, headers.indexOf("Date") + 1).setValue(DateStr);
        if (Payment_Method && headers.indexOf("Payment_Method") !== -1) {
          sheet.getRange(rowNum, headers.indexOf("Payment_Method") + 1).setValue(Payment_Method);
        }
        updated = true;
        break;
      }
    }
    
    if (updated) {
      SpreadsheetApp.flush();
      registrarLogAuditoria(userEmail || "Admin", "EDITAR", "MOVIMIENTO", `Modificado movimiento ID ${Movimiento_ID}: "${Concept}" por $${parseFloat(Amount).toLocaleString('es-AR')}`);
      return { success: true, message: "Movimiento actualizado correctamente." };
    } else {
      return { success: false, message: "No se encontró el movimiento a modificar." };
    }
  } catch (error) {
    console.error("Error en editarMovimientoTorneo:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Elimina un movimiento financiero de caja.
 */
function eliminarMovimientoTorneo(movId, userEmail) {
  try {
    if (!movId) throw new Error("ID de movimiento no especificado.");
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_FINANZAS);
    if (!sheet) throw new Error("No se encontró la hoja de finanzas.");
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const movIdCol = headers.indexOf("Movimiento_ID");
    
    let deleted = false;
    let conceptDeleted = "";
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][movIdCol].toString() === movId.toString()) {
        conceptDeleted = data[i][headers.indexOf("Concept")];
        sheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }
    
    if (deleted) {
      SpreadsheetApp.flush();
      registrarLogAuditoria(userEmail || "Admin", "ELIMINAR", "MOVIMIENTO", `Eliminado movimiento ID ${movId}: "${conceptDeleted}"`);
      return { success: true, message: "Movimiento eliminado de la caja." };
    } else {
      return { success: false, message: "No se encontró el movimiento a eliminar." };
    }
  } catch (error) {
    console.error("Error en eliminarMovimientoTorneo:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Obtiene el listado completo de logs de auditoría para administradores.
 */
function obtenerLogsAuditoria() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(HOJA_LOGS);
    if (!sheet) {
      return { success: true, logs: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, logs: [] };
    
    const headers = data[0];
    const rows = data.slice(1);
    
    const logs = rows.map(row => {
      const item = {};
      headers.forEach((h, idx) => item[h] = row[idx]);
      return item;
    });
    
    logs.sort((a, b) => b.Timestamp.localeCompare(a.Timestamp));
    return { success: true, logs: logs };
  } catch (error) {
    console.error("Error en obtenerLogsAuditoria:", error);
    return { success: false, message: error.message, logs: [] };
  }
}

/**
 * Programa o registra un nuevo partido en la agenda deportiva.
 */
function registrarPartidoNuevo(partidoObj) {
  try {
    const { Torneo_ID, DateStr, Opponent, Location } = partidoObj;
    if (!Torneo_ID || !DateStr || !Opponent || !Location) {
      throw new Error("Faltan parámetros obligatorios para programar el partido.");
    }
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_PARTIDOS);
    if (!sheet) throw new Error("No se encontró la hoja de partidos.");
    
    const partidoId = `PAR-${Date.now().toString().slice(-6)}`;
    
    // Columnas: Partido_ID, Torneo_ID, Date, Opponent, Location, Result
    sheet.appendRow([partidoId, Torneo_ID, DateStr, Opponent, Location, "Pendiente"]);
    SpreadsheetApp.flush();
    
    return { success: true, message: "Partido programado exitosamente." };
    
  } catch (error) {
    console.error("Error en registrarPartidoNuevo:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Genera el link de Mercado Pago para un pago específico si no existe, o retorna el ya existente.
 */
function generarPagoMercadoPago(paymentId) {
  try {
    const ss = getSpreadsheet();
    const sheetPagos = ss.getSheetByName(HOJA_PAGOS);
    const pagosData = sheetPagos.getDataRange().getValues();
    const headers = pagosData[0];
    
    const idColIndex = headers.indexOf("Payment_ID");
    const linkColIndex = headers.indexOf("MP_Link");
    let emailColIndex = headers.indexOf("Email");
    if (emailColIndex === -1) emailColIndex = headers.indexOf("User_Email");
    
    const amountColIndex = headers.indexOf("Amount");
    const monthColIndex = headers.indexOf("Month");
    const statusColIndex = headers.indexOf("Status");
    
    let filaIndex = -1;
    for (let i = 1; i < pagosData.length; i++) {
      if (pagosData[i][idColIndex].toString().trim() === paymentId.toString().trim()) {
        filaIndex = i;
        break;
      }
    }
    
    if (filaIndex === -1) throw new Error("No se encontró el registro de pago.");
    if (pagosData[filaIndex][statusColIndex] === "Pagado") {
      throw new Error("Este pago ya ha sido registrado como Pagado.");
    }
    
    const linkExistente = pagosData[filaIndex][linkColIndex];
    if (linkExistente && linkExistente !== "") {
      return { success: true, mp_link: linkExistente };
    }
    
    const email = pagosData[filaIndex][emailColIndex];
    const amount = parseFloat(pagosData[filaIndex][amountColIndex]);
    const month = pagosData[filaIndex][monthColIndex];
    
    const mpLink = crearPreferenciaMercadoPago({
      paymentId,
      email,
      amount,
      month
    });
    
    sheetPagos.getRange(filaIndex + 1, linkColIndex + 1).setValue(mpLink);
    SpreadsheetApp.flush();
    
    return { success: true, mp_link: mpLink };
    
  } catch (error) {
    console.error("Error en generarPagoMercadoPago:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Permite al administrador o Coach registrar el cobro en efectivo de una cuota.
 */
function marcarPagoComoPagado(paymentId, collectorEmail, collectorRole) {
  try {
    const ss = getSpreadsheet();
    const sheetPagos = ss.getSheetByName(HOJA_PAGOS);
    const pagosData = sheetPagos.getDataRange().getValues();
    const headers = pagosData[0];
    
    const idColIndex = headers.indexOf("Payment_ID");
    const statusColIndex = headers.indexOf("Status");
    let byColIndex = headers.indexOf("Collected_By");
    let atColIndex = headers.indexOf("Collected_At");
    
    if (byColIndex === -1) byColIndex = 5;
    if (atColIndex === -1) atColIndex = 6;
    
    let filaIndex = -1;
    for (let i = 1; i < pagosData.length; i++) {
      if (pagosData[i][idColIndex].toString().trim() === paymentId.toString().trim()) {
        filaIndex = i;
        break;
      }
    }
    
    if (filaIndex === -1) throw new Error("No se encontró el registro de pago.");
    
    const rowNum = filaIndex + 1;
    const nowStr = new Date().toISOString().replace("T", " ").substring(0, 16);
    const collectorStr = collectorRole || 'Admin';
    
    sheetPagos.getRange(rowNum, statusColIndex + 1).setValue("Pagado");
    sheetPagos.getRange(rowNum, byColIndex + 1).setValue(collectorStr);
    sheetPagos.getRange(rowNum, atColIndex + 1).setValue(nowStr);
    SpreadsheetApp.flush();
    
    registrarLogAuditoria(collectorEmail || "Admin", "MODIFICAR", "PAGO", `Cobro en efectivo registrado por ${collectorStr} (${collectorEmail || 'Admin'}) para pago ID ${paymentId} el ${nowStr}`);
    
    return { success: true, message: "El pago fue registrado en efectivo exitosamente." };
    
  } catch (error) {
    console.error("Error en marcarPagoComoPagado:", error);
    return { success: false, message: error.message };
  }
}

// ==========================================
// FUNCIONES AUXILIARES INTERNAS
// ==========================================

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function crearPreferenciaMercadoPago({ paymentId, email, amount, month }) {
  const token = getMercadoPagoToken();
  const url = "https://api.mercadopago.com/checkout/preferences";
  const numericAmount = parseFloat(amount || 0);

  const body = {
    items: [
      {
        id: String(paymentId),
        title: `Futsal Haedo - Cuota Mes: ${month}`,
        description: `Pago de arancel mensual - Socio: ${email}`,
        quantity: 1,
        unit_price: numericAmount,
        currency_id: "ARS"
      }
    ],
    external_reference: String(paymentId),
    back_urls: {
      success: "https://haedofutsal-app.onrender.com",
      pending: "https://haedofutsal-app.onrender.com",
      failure: "https://haedofutsal-app.onrender.com"
    },
    auto_return: "approved"
  };

  // Incluir payer solo si es un correo real que no sea de prueba interna
  if (email && email.includes("@") && !email.includes("futsalhaedo.com")) {
    body.payer = { email: email };
  }

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": `Bearer ${token}`
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode < 200 || responseCode >= 300) {
    throw new Error(`Error en API Mercado Pago (HTTP ${responseCode}): ${responseText}`);
  }

  const responseData = JSON.parse(responseText);

  if (responseData) {
    // Si la credencial es de Pruebas (TEST-...), Mercado Pago requiere sandbox_init_point
    if (token.startsWith("TEST-") && responseData.sandbox_init_point) {
      return responseData.sandbox_init_point;
    }
    if (responseData.init_point) {
      return responseData.init_point;
    }
  }
  
  throw new Error("La respuesta de Mercado Pago no contiene un enlace de checkout válido.");
}

/**
 * Crea las hojas de cálculo, escribe las cabeceras de columnas y carga datos iniciales
 * de prueba ampliados (con Torneos, Finanzas y Partidos) para Futsal Haedo.
 */
function inicializarBaseDatos() {
  const ss = getSpreadsheet();
  
  // 1. Configurar Hoja "Admins"
  let sheetAdmins = ss.getSheetByName(HOJA_ADMINS);
  if (!sheetAdmins) {
    sheetAdmins = ss.insertSheet(HOJA_ADMINS);
  }
  sheetAdmins.clear();
  sheetAdmins.appendRow(["Email"]);
  sheetAdmins.appendRow(["admin@futsalhaedo.com"]);
  sheetAdmins.appendRow(["guillermopazos@gmail.com"]);
  
  // 2. Configurar Hoja "Categorias"
  let sheetCats = ss.getSheetByName(HOJA_CATEGORIAS);
  if (!sheetCats) {
    sheetCats = ss.insertSheet(HOJA_CATEGORIAS);
  }
  sheetCats.clear();
  sheetCats.appendRow(["Category_ID", "Name", "Coach", "Monthly_Fee", "Torneos"]);
  sheetCats.appendRow(["Fut-May+35", "Futsal Mayores +35", "Bocha", 15000, "EDEFI"]);
  sheetCats.appendRow(["Fut-May+42", "Futsal Mayores +42", "Bocha", 15000, "EDEFI"]);
  sheetCats.appendRow(["Fut-Fem-1ra", "Futsal Femenino 1ra", "Pol", 20000, "FutFEM"]);
  sheetCats.appendRow(["Fut-Juv-5ta", "Futsal Juveniles 5ta", "Bocha", 30000, "FUTSALA"]);
  sheetCats.appendRow(["Bab-2015", "Baby Futbol - 2015", "Marcelo", 30000, "Baby EDEFI"]);
  sheetCats.getRange("D2:D6").setNumberFormat("$#,##0.00");
  
  // 3. Configurar Hoja "Usuarios"
  let sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(HOJA_USUARIOS);
  }
  sheetUsers.clear();
  sheetUsers.appendRow(["Email", "Role", "Name", "Phone", "Category"]);
  sheetUsers.appendRow(["admin@futsalhaedo.com", "Admin", "Administrador Haedo", "1122334455", ""]);
  sheetUsers.appendRow(["guillermopazos@gmail.com", "Admin", "Guillermo Pazos", "1133445566", ""]);
  sheetUsers.appendRow(["deportista@futsalhaedo.com", "Deportista", "Juan Perez", "1199887766", "Fut-May+35"]);
  
  // Socios de prueba (20)
  const socios = [
    ["socio1@futsalhaedo.com", "Deportista", "Carlos Bianchi", "1101000000", "Fut-May+35"],
    ["socio2@futsalhaedo.com", "Deportista", "Diego Maradona", "1102000000", "Fut-May+35"],
    ["socio3@futsalhaedo.com", "Deportista", "Lionel Messi", "1103000000", "Fut-May+35"],
    ["socio4@futsalhaedo.com", "Deportista", "Roman Riquelme", "1104000000", "Fut-May+35"],
    ["socio5@futsalhaedo.com", "Deportista", "Enzo Francescoli", "1105000000", "Fut-May+42"],
    ["socio6@futsalhaedo.com", "Deportista", "Martin Palermo", "1106000000", "Fut-May+42"],
    ["socio7@futsalhaedo.com", "Deportista", "Ariel Ortega", "1107000000", "Fut-May+42"],
    ["socio8@futsalhaedo.com", "Deportista", "Javier Zanetti", "1108000000", "Fut-May+42"],
    ["socio9@futsalhaedo.com", "Deportista", "Estefania Banini", "1109000000", "Fut-Fem-1ra"],
    ["socio10@futsalhaedo.com", "Deportista", "Yamila Rodriguez", "1110000000", "Fut-Fem-1ra"],
    ["socio11@futsalhaedo.com", "Deportista", "Florencia Bonsegundo", "1111000000", "Fut-Fem-1ra"],
    ["socio12@futsalhaedo.com", "Deportista", "Vanina Correa", "1112000000", "Fut-Fem-1ra"],
    ["socio13@futsalhaedo.com", "Deportista", "Julian Alvarez", "1113000000", "Fut-Juv-5ta"],
    ["socio14@futsalhaedo.com", "Deportista", "Enzo Fernandez", "1114000000", "Fut-Juv-5ta"],
    ["socio15@futsalhaedo.com", "Deportista", "Alexis Mac Allister", "1115000000", "Fut-Juv-5ta"],
    ["socio16@futsalhaedo.com", "Deportista", "Alejandro Garnacho", "1116000000", "Fut-Juv-5ta"],
    ["socio17@futsalhaedo.com", "Deportista", "Thiago Messi", "1117000000", "Bab-2015"],
    ["socio18@futsalhaedo.com", "Deportista", "Benjamin Aguero", "1118000000", "Bab-2015"],
    ["socio19@futsalhaedo.com", "Deportista", "Mateo Ronaldo", "1119000000", "Bab-2015"],
    ["socio20@futsalhaedo.com", "Deportista", "Ciro Messi", "1120000000", "Bab-2015"]
  ];
  socios.forEach(row => sheetUsers.appendRow(row));
  
  // 4. Configurar Hoja "Pagos"
  let sheetPagos = ss.getSheetByName(HOJA_PAGOS);
  if (!sheetPagos) {
    sheetPagos = ss.insertSheet(HOJA_PAGOS);
  }
  sheetPagos.clear();
  sheetPagos.appendRow(["Payment_ID", "Email", "Month", "Amount", "Status", "MP_Link"]);
  
  const aranceles = {
    "Fut-May+35": 15000,
    "Fut-May+42": 15000,
    "Fut-Fem-1ra": 20000,
    "Fut-Juv-5ta": 30000,
    "Bab-2015": 30000
  };
  
  const morosos = {
    "deportista@futsalhaedo.com": ["2026-06"],
    "socio3@futsalhaedo.com": ["2026-06"],
    "socio7@futsalhaedo.com": ["2026-05", "2026-06"],
    "socio12@futsalhaedo.com": ["2026-06"],
    "socio16@futsalhaedo.com": ["2026-06"],
    "socio20@futsalhaedo.com": ["2026-05", "2026-06"]
  };
  
  const todosLosDeportistas = [["deportista@futsalhaedo.com", "Deportista", "Juan Perez", "1199887766", "Fut-May+35"], ...socios];
  let pCounter = 1;
  todosLosDeportistas.forEach(d => {
    const email = d[0];
    const cat = d[4];
    const fee = aranceles[cat] || 10000;
    
    ["2026-05", "2026-06"].forEach(month => {
      const isMoroso = morosos[email] && morosos[email].includes(month);
      const pid = `PAG-${pCounter.toString().padStart(4, '0')}`;
      pCounter++;
      
      sheetPagos.appendRow([
        pid,
        email,
        month,
        fee,
        isMoroso ? "Pendiente" : "Pagado",
        isMoroso ? "" : "https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_" + pid
      ]);
    });
  });
  
  const numPagos = sheetPagos.getLastRow();
  if (numPagos > 1) {
    sheetPagos.getRange(2, 4, numPagos - 1, 1).setNumberFormat("$#,##0.00");
  }
  
  // 5. Configurar Hoja "Torneos"
  let sheetTorneos = ss.getSheetByName(HOJA_TORNEOS);
  if (!sheetTorneos) {
    sheetTorneos = ss.insertSheet(HOJA_TORNEOS);
  }
  sheetTorneos.clear();
  sheetTorneos.appendRow(["Torneo_ID", "Name", "Category"]);
  sheetTorneos.appendRow(["T-001", "Torneo EDEFI +35", "Fut-May+35"]);
  sheetTorneos.appendRow(["T-002", "Torneo EDEFI +42", "Fut-May+42"]);
  sheetTorneos.appendRow(["T-003", "Liga Femenina FutFEM", "Fut-Fem-1ra"]);
  sheetTorneos.appendRow(["T-004", "Liga Juvenil FUTSALA", "Fut-Juv-5ta"]);
  sheetTorneos.appendRow(["T-005", "Liga Baby EDEFI", "Bab-2015"]);
  
  // 6. Configurar Hoja "Finanzas_Torneos"
  let sheetFinanzas = ss.getSheetByName(HOJA_FINANZAS);
  if (!sheetFinanzas) {
    sheetFinanzas = ss.insertSheet(HOJA_FINANZAS);
  }
  sheetFinanzas.clear();
  sheetFinanzas.appendRow(["Movimiento_ID", "Torneo_ID", "Type", "Concept", "Amount", "Date"]);
  
  // 7. Configurar Hoja "Partidos"
  let sheetPartidos = ss.getSheetByName(HOJA_PARTIDOS);
  if (!sheetPartidos) {
    sheetPartidos = ss.insertSheet(HOJA_PARTIDOS);
  }
  sheetPartidos.clear();
  sheetPartidos.appendRow(["Partido_ID", "Torneo_ID", "Date", "Opponent", "Location", "Result"]);

  const fechas = ["2026-06-06", "2026-06-13", "2026-06-20", "2026-06-27", "2026-07-04"];
  const rivales = [
    ["Moron Futsal", "Ateneo Haedo", "Estudiantil Porteño", "Club Leloir", "Nito Futsal"],
    ["Defensores de Haedo", "Ramos Mejía", "Huracán San Justo", "Deportivo Morón", "UAI Urquiza"],
    ["Social Club Fem", "Sportivo Haedo", "Ituzaingó Fem", "La Matanza", "Almagro Femenino"],
    ["Villa Pearson Juv", "CIDECO Juv", "Nueva Estrella", "Club Leloir Juv", "Defensores Juv"],
    ["Lomas Baby", "Almafuerte Baby", "Parque Baby", "Social Haedo Baby", "Gimnasia Baby"]
  ];

  const torneosIds = ["T-001", "T-002", "T-003", "T-004", "T-005"];
  let partidoIdCounter = 1;
  let movimientoIdCounter = 1;

  torneosIds.forEach((tId, tIdx) => {
    fechas.forEach((fecha, fIdx) => {
      const opponent = rivales[tIdx][fIdx];
      const isLocal = fIdx % 2 === 0;
      
      let result = "Pendiente";
      if (fIdx === 0) result = (tIdx % 2 === 0) ? "4 - 2" : "2 - 0";
      else if (fIdx === 1) result = (tIdx % 2 === 0) ? "1 - 1" : "0 - 3";
      else if (fIdx === 2) result = (tIdx % 2 === 0) ? "3 - 2" : "2 - 2";
      else if (fIdx === 3 && tIdx === 0) result = "2 - 1";
      
      const pId = `PAR-${partidoIdCounter.toString().padStart(4, '0')}`;
      partidoIdCounter++;
      
      sheetPartidos.appendRow([pId, tId, fecha, opponent, isLocal ? "Local" : "Visitante", result]);
      
      if (isLocal) {
        const alquilerMonto = 8000 + (tIdx * 1000);
        const arbitrajeMonto = 5000 + (tIdx * 500);
        const entradasMonto = 12000 + (fIdx * 1500) + (tIdx * 1000);
        const buffetMonto = 6000 + (fIdx * 1000) + (tIdx * 2000);
        
        const movId1 = `MOV-${movimientoIdCounter.toString().padStart(4, '0')}`;
        movimientoIdCounter++;
        sheetFinanzas.appendRow([movId1, tId, "Gasto", `Alquiler Cancha - Fecha ${fIdx + 1} vs ${opponent}`, alquilerMonto, fecha]);
        
        const movId2 = `MOV-${movimientoIdCounter.toString().padStart(4, '0')}`;
        movimientoIdCounter++;
        sheetFinanzas.appendRow([movId2, tId, "Gasto", `Arbitraje - Fecha ${fIdx + 1} vs ${opponent}`, arbitrajeMonto, fecha]);
        
        const movId3 = `MOV-${movimientoIdCounter.toString().padStart(4, '0')}`;
        movimientoIdCounter++;
        sheetFinanzas.appendRow([movId3, tId, "Ingreso", `Venta Entradas - Fecha ${fIdx + 1} vs ${opponent}`, entradasMonto, fecha]);
        
        const movId4 = `MOV-${movimientoIdCounter.toString().padStart(4, '0')}`;
        movimientoIdCounter++;
        sheetFinanzas.appendRow([movId4, tId, "Ingreso", `Buffet/Cantina - Fecha ${fIdx + 1} vs ${opponent}`, buffetMonto, fecha]);
      } else {
        const viaticoMonto = 4000 + (tIdx * 1000);
        const rifasMonto = 5000 + (fIdx * 1200) + (tIdx * 1500);
        
        const movId1 = `MOV-${movimientoIdCounter.toString().padStart(4, '0')}`;
        movimientoIdCounter++;
        sheetFinanzas.appendRow([movId1, tId, "Gasto", `Transporte/Viáticos - Fecha ${fIdx + 1} vs ${opponent}`, viaticoMonto, fecha]);
        
        const movId2 = `MOV-${movimientoIdCounter.toString().padStart(4, '0')}`;
        movimientoIdCounter++;
        sheetFinanzas.appendRow([movId2, tId, "Ingreso", `Venta de Rifas - Fecha ${fIdx + 1} vs ${opponent}`, rifasMonto, fecha]);
      }
    });
  });
  
  const numMovs = sheetFinanzas.getLastRow();
  if (numMovs > 1) {
    sheetFinanzas.getRange(2, 5, numMovs - 1, 1).setNumberFormat("$#,##0.00");
  }
  
  // Limpiar la hoja por defecto
  const defaultSheet = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
  
  console.log("Inicialización ampliada completada. Hojas deportivas y financieras configuradas con datos.");
}
