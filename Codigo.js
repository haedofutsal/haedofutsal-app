/**
 * Google Apps Script - Backend API para ERP Web Haedo Futsal (Ampliando Torneos, Finanzas y Socios)
 * Desarrollado con sintaxis moderna ES6 (motor V8).
 * 
 * Este archivo actúa como el backend que sirve el frontend HTML (HtmlService)
 * y expone las funciones del servidor llamadas de forma segura desde el cliente
 * mediante "google.script.run".
 */

// ==========================================
// CONFIGURACIÓN Y CONSTANTES
// ==========================================
const SPREADSHEET_ID = "1nDZWIxdGPWK8YY86JcOW9YCwTuApQX380H7Wk4JGjWg"; // Planilla oficial de Haedo Futsal
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
    if (e && e.parameter && e.parameter.api === "run") {
      const functionName = e.parameter.func;
      const args = e.parameter.args ? JSON.parse(decodeURIComponent(e.parameter.args)) : [];
      if (typeof this[functionName] === 'function') {
        const result = this[functionName].apply(this, args);
        return ContentService.createTextOutput(JSON.stringify({ success: true, result: result }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Función no encontrada: " + functionName }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    const template = HtmlService.createTemplateFromFile("Index");
    return template.evaluate()
      .setTitle("Haedo Futsal - ERP Deportivo")
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
function loginUsuario(username, password) {
  try {
    if (!username) throw new Error("El usuario es requerido.");
    const userNorm = username.toLowerCase().trim();
    const ss = getSpreadsheet();
    
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheetUsers) throw new Error("La tabla de usuarios no está inicializada.");
    
    const usersData = sheetUsers.getDataRange().getValues();
    const headers = usersData[0];
    
    const userColIndex = headers.indexOf("Username") !== -1 ? headers.indexOf("Username") : headers.indexOf("Email");
    const passColIndex = headers.indexOf("Password");
    
    if (userColIndex === -1) throw new Error("No se encontró la columna de usuario en el sistema.");
    
    let userRow = null;
    for (let i = 1; i < usersData.length; i++) {
      const dbUser = (usersData[i][userColIndex] || "").toString().toLowerCase().trim();
      const dbEmail = headers.indexOf("Email") !== -1 ? (usersData[i][headers.indexOf("Email")] || "").toString().toLowerCase().trim() : "";
      if (dbUser === userNorm || (dbEmail && dbEmail === userNorm)) {
        userRow = usersData[i];
        break;
      }
    }
    
    if (!userRow) {
      return { success: false, message: "El usuario no se encuentra registrado en el sistema." };
    }
    
    const userData = {};
    headers.forEach((header, index) => {
      userData[header] = userRow[index];
    });
    
    // Si viene la clave, verificarla (si no viene, es autologin por sesión activa)
    if (password !== null && password !== undefined) {
      const dbPass = passColIndex !== -1 ? (userData.Password || "").toString().trim() : "1234";
      if (dbPass !== password.toString().trim()) {
        return { success: false, message: "La clave ingresada es incorrecta." };
      }
    }
    
    return {
      success: true,
      role: userData.Role || "Deportista",
      email: userData.Email || "",
      name: userData.Name || "Socio",
      username: userData.Username || username,
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
    
    // 3. Obtener partidos programados para todas las categorías del deportista
    const misPartidos = [];
    const sheetTorneos = ss.getSheetByName(HOJA_TORNEOS);
    const sheetPartidos = ss.getSheetByName(HOJA_PARTIDOS);

    // Soportar múltiples categorías separadas por |
    const socioCategories = socioCategory
      ? socioCategory.split('|').map(c => c.trim()).filter(Boolean)
      : [];

    if (sheetTorneos && sheetPartidos && socioCategories.length > 0) {
      const torneosData = sheetTorneos.getDataRange().getValues();
      const tHeaders = torneosData[0];
      const tIdCol = tHeaders.indexOf("Torneo_ID");
      const tCatCol = tHeaders.indexOf("Category");
      const tNameCol = tHeaders.indexOf("Name");

      // Obtener torneos que disputa cualquiera de sus categorías
      const misTorneosIds = [];
      const torneosMap = {};
      for (let i = 1; i < torneosData.length; i++) {
        const tCat = getVal(torneosData[i], tCatCol);
        if (socioCategories.includes(tCat)) {
          const tId = getVal(torneosData[i], tIdCol);
          if (!misTorneosIds.includes(tId)) misTorneosIds.push(tId);
          torneosMap[tId] = {
            name: getVal(torneosData[i], tNameCol),
            category: tCat
          };
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
          partidoObj.Torneo_Name = torneosMap[torneoId] ? torneosMap[torneoId].name : "Torneo";
          partidoObj.Category = torneosMap[torneoId] ? torneosMap[torneoId].category : "";
          misPartidos.push(partidoObj);
        }
      }
    }
    
    // Cargar nombres de categorías
    const sheetCategorias = ss.getSheetByName(HOJA_CATEGORIAS);
    const categoriasMap = {};
    if (sheetCategorias) {
      const catData = sheetCategorias.getDataRange().getValues();
      const catHeaders = catData[0];
      const catIdCol = catHeaders.indexOf("Category_ID");
      const catNameCol = catHeaders.indexOf("Name");
      for (let i = 1; i < catData.length; i++) {
        const cId = getVal(catData[i], catIdCol);
        if (cId) {
          categoriasMap[cId] = getVal(catData[i], catNameCol);
        }
      }
    }
    
    // Ordenar partidos por fecha
    misPartidos.sort((a, b) => (b.Date || "").localeCompare(a.Date || ""));
    
    return {
      success: true,
      perfil,
      pagos: misPagos,
      partidos: misPartidos,
      categoriasMap: categoriasMap
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
    const tTicketCol = tHeaders.indexOf("Ticket_Price");
    
    rowsTorneos.forEach(row => {
      const tId = getVal(row, tIdCol);
      if (!tId) return;
      torneos.push({
        Torneo_ID: tId,
        Name: getVal(row, tNameCol),
        Category: getVal(row, tCatCol),
        Ticket_Price: parseFloat(getVal(row, tTicketCol)) || 0,
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
    
    // 5. Cargar Categorías
    const categorias = [];
    const sheetCategorias = ss.getSheetByName(HOJA_CATEGORIAS);
    if (sheetCategorias) {
      const catData = sheetCategorias.getDataRange().getValues();
      const catHeaders = catData[0] || [];
      const rowsCategorias = catData.slice(1);
      
      const catIdCol = catHeaders.indexOf("Category_ID");
      const catNameCol = catHeaders.indexOf("Name");
      const catCoachCol = catHeaders.indexOf("Coach");
      const catFeeCol = catHeaders.indexOf("Monthly_Fee");
      const catTorneosCol = catHeaders.indexOf("Torneos");
      
      rowsCategorias.forEach(row => {
        const cId = getVal(row, catIdCol);
        if (!cId) return;
        categorias.push({
          Category_ID: cId,
          Name: getVal(row, catNameCol),
          Coach: getVal(row, catCoachCol),
          Monthly_Fee: parseFloat(getVal(row, catFeeCol)) || 0,
          Torneos: getVal(row, catTorneosCol)
        });
      });
    }
    
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
      categorias: categorias,
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

    const username = (socioObj.Username || "").trim().toLowerCase();
    if (!username) throw new Error("El nombre de usuario es obligatorio.");
    
    const password = (socioObj.Password || "").toString().trim();
    if (password.length !== 4 || isNaN(password)) throw new Error("La clave debe tener 4 dígitos numéricos.");

    // Verificar si ya existe email o username
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailColIdx = headers.indexOf("Email");
    const userColIdx = headers.indexOf("Username");

    for (let i = 1; i < data.length; i++) {
      if (emailColIdx !== -1 && (data[i][emailColIdx] || "").toString().toLowerCase().trim() === email) {
        throw new Error("Ya existe un deportista registrado con el correo " + email);
      }
      if (userColIdx !== -1 && (data[i][userColIdx] || "").toString().toLowerCase().trim() === username) {
        throw new Error("Ya existe un usuario registrado con el nombre de usuario " + username);
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

    // Armar el array dinámicamente según cabeceras
    const rowArray = [];
    headers.forEach(h => {
      if (h === "Email") rowArray.push(email);
      else if (h === "Role") rowArray.push(assignedRole);
      else if (h === "Name") rowArray.push(socioObj.Name || "");
      else if (h === "Phone") rowArray.push(socioObj.Phone || "");
      else if (h === "Category") rowArray.push(socioObj.Category || "Fut-May+35");
      else if (h === "DNI") rowArray.push(socioObj.DNI || "");
      else if (h === "BirthDate") rowArray.push(socioObj.BirthDate || "");
      else if (h === "Age") rowArray.push(age ? age.toString() : "");
      else if (h === "JoinDate") rowArray.push(todayStr);
      else if (h === "BloodType") rowArray.push(socioObj.BloodType || "O+");
      else if (h === "MedicalFit") rowArray.push(socioObj.MedicalFit || "Apto Físico Vigente");
      else if (h === "ObraSocial") rowArray.push(socioObj.ObraSocial || "Particular");
      else if (h === "EmergencyContact") rowArray.push(socioObj.EmergencyContact || socioObj.ParentName || "");
      else if (h === "EmergencyPhone") rowArray.push(socioObj.EmergencyPhone || socioObj.ParentPhone || "");
      else if (h === "ParentName") rowArray.push(isMinor ? (socioObj.ParentName || "") : "");
      else if (h === "ParentPhone") rowArray.push(isMinor ? (socioObj.ParentPhone || "") : "");
      else if (h === "Notes") rowArray.push(socioObj.Notes || "");
      else if (h === "Photo") rowArray.push("");
      else if (h === "Username") rowArray.push(username);
      else if (h === "Password") rowArray.push(password);
      else rowArray.push("");
    });

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

/**
 * Registra el pago de una cuota por transferencia tras recibir y validar un comprobante.
 * Cambia el estado del pago a 'Pagado' y suma el importe a las finanzas del club.
 */
function registrarPagoTransferenciaComprobante(paymentId, email, amount, month, paymentMethod) {
  try {
    const ss = getSpreadsheet();
    
    // 1. Marcar pago como Pagado
    const sheetPagos = ss.getSheetByName(HOJA_PAGOS);
    if (!sheetPagos) throw new Error("No se encontró la hoja de pagos.");
    const pagosData = sheetPagos.getDataRange().getValues();
    const pHeaders = pagosData[0];
    const pIdCol = pHeaders.indexOf("Payment_ID");
    const pStatusCol = pHeaders.indexOf("Status");
    const pByCol = pHeaders.indexOf("Collected_By");
    const pAtCol = pHeaders.indexOf("Collected_At");
    const pEmailCol = pHeaders.indexOf("Email");
    
    let filaIndex = -1;
    for (let i = 1; i < pagosData.length; i++) {
      if (pagosData[i][pIdCol].toString().trim() === paymentId.toString().trim()) {
        filaIndex = i;
        break;
      }
    }
    
    if (filaIndex === -1) {
      throw new Error("No se encontró el registro de pago.");
    }
    
    const rowNum = filaIndex + 1;
    const nowStr = new Date().toISOString().replace("T", " ").substring(0, 16);
    const methodStr = paymentMethod || "Transferencia MP";
    const socioEmail = pagosData[filaIndex][pEmailCol] || email;
    
    sheetPagos.getRange(rowNum, pStatusCol + 1).setValue("Pagado");
    sheetPagos.getRange(rowNum, pByCol + 1).setValue(`Socio (${methodStr})`);
    sheetPagos.getRange(rowNum, pAtCol + 1).setValue(nowStr);
    
    // 2. Sumar el importe a las finanzas del club (appendRow en Finanzas_Torneos)
    const sheetFinanzas = ss.getSheetByName(HOJA_FINANZAS);
    if (sheetFinanzas) {
      const movId = `MOV-${Date.now().toString().slice(-6)}`;
      const dateFormatted = new Date().toISOString().split("T")[0]; // YYYY-MM-DD
      const concept = `Cuota Social ${month} - Socio: ${socioEmail}`;
      const numericAmount = parseFloat(amount || 0);
      
      // Columnas: Movimiento_ID, Torneo_ID, Type, Concept, Amount, Date, Payment_Method
      sheetFinanzas.appendRow([movId, "General", "Ingreso", concept, numericAmount, dateFormatted, methodStr]);
      
      const lastRowIndex = sheetFinanzas.getLastRow();
      if (sheetFinanzas.getRange(lastRowIndex, 5).setNumberFormat) {
        sheetFinanzas.getRange(lastRowIndex, 5).setNumberFormat("$#,##0.00");
      }
    }
    
    SpreadsheetApp.flush();
    
    // 3. Registrar Log
    registrarLogAuditoria(socioEmail, "MODIFICAR", "PAGO", `Pago por transferencia autogestionado para cuota ${month} por $${parseFloat(amount).toLocaleString('es-AR')} (Comprobante recibido y validado)`);
    
    return { success: true, message: "Comprobante validado correctamente. Tu cuota ha sido acreditada y registrada en las finanzas." };
    
  } catch (error) {
    console.error("Error en registrarPagoTransferenciaComprobante:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Obtiene los datos personales de un socio por su Email, DNI, Teléfono o Username.
 */
function obtenerDatosSocioPublico(identifier) {
  try {
    const ss = getSpreadsheet();
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheetUsers) throw new Error("No se encontró la hoja de usuarios.");

    const usersData = sheetUsers.getDataRange().getValues();
    const userHeaders = usersData[0];
    const emailColIndex = userHeaders.indexOf("Email");
    const dniColIndex = userHeaders.indexOf("DNI");
    const phoneColIndex = userHeaders.indexOf("Phone");
    const userColIndex = userHeaders.indexOf("Username");
    
    if (emailColIndex === -1) {
      throw new Error("Estructura de la hoja de usuarios inválida.");
    }

    const getVal = (row, idx) => (row && idx !== -1 && idx < row.length && row[idx] !== undefined && row[idx] !== null) ? row[idx].toString().trim() : "";

    const searchVal = (identifier || "").toString().toLowerCase().trim();
    // Remover caracteres no alfanuméricos para búsquedas limpias de DNI o Teléfono
    const cleanSearchVal = searchVal.replace(/[^a-z0-9]/g, "");

    let perfil = null;
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      const email = getVal(row, emailColIndex).toLowerCase();
      const dni = getVal(row, dniColIndex).replace(/[^0-9]/g, "");
      const phone = getVal(row, phoneColIndex).replace(/[^0-9]/g, "");
      const username = userColIndex !== -1 ? getVal(row, userColIndex).toLowerCase() : "";

      if (
        email === searchVal ||
        (dni && dni === cleanSearchVal) ||
        (phone && (phone === cleanSearchVal || phone.endsWith(cleanSearchVal) || cleanSearchVal.endsWith(phone))) ||
        (username && username === searchVal)
      ) {
        perfil = {};
        userHeaders.forEach((header, index) => {
          perfil[header] = getVal(row, index);
        });
        break;
      }
    }

    if (!perfil) {
      throw new Error("No se encontró ningún socio registrado con el identificador provisto.");
    }

    return {
      success: true,
      datos: {
        Email: perfil.Email || "",
        Name: perfil.Name || "",
        Phone: perfil.Phone || "",
        DNI: perfil.DNI || "",
        BirthDate: perfil.BirthDate || "",
        BloodType: perfil.BloodType || "",
        ObraSocial: perfil.ObraSocial || "",
        MedicalFit: perfil.MedicalFit || "",
        EmergencyContact: perfil.EmergencyContact || "",
        EmergencyPhone: perfil.EmergencyPhone || ""
      }
    };
  } catch (error) {
    console.error("Error en obtenerDatosSocioPublico:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Actualiza los datos de un socio desde el formulario público de WhatsApp.
 */
function actualizarDatosSocioPublico(email, datos) {
  try {
    const ss = getSpreadsheet();
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheetUsers) throw new Error("No se encontró la hoja de usuarios.");

    const emailClean = (email || "").trim().toLowerCase();
    const usersData = sheetUsers.getDataRange().getValues();
    const userHeaders = usersData[0];
    const emailColIndex = userHeaders.indexOf("Email");

    if (emailColIndex === -1) {
      throw new Error("Estructura de la hoja de usuarios inválida.");
    }

    let foundRowIdx = -1;
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      if ((row[emailColIndex] || "").toString().toLowerCase().trim() === emailClean) {
        foundRowIdx = i + 1; // +1 for headers
        break;
      }
    }

    if (foundRowIdx === -1) {
      throw new Error("No se encontró ningún socio registrado con el correo especificado.");
    }

    // Capitalizar nombre y apellido si vienen presentes
    let formattedName = datos.Name || "";
    if (formattedName) {
      formattedName = formattedName.trim().split(/\s+/).map(w => {
        if (!w) return "";
        return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
      }).join(" ");
    }

    // Actualizar columnas una por una
    const updates = {
      "Name": formattedName,
      "Phone": (datos.Phone || "").trim(),
      "DNI": (datos.DNI || "").trim(),
      "BirthDate": (datos.BirthDate || "").trim(),
      "BloodType": (datos.BloodType || "").trim(),
      "ObraSocial": (datos.ObraSocial || "").trim(),
      "MedicalFit": (datos.MedicalFit || "").trim(),
      "EmergencyContact": (datos.EmergencyContact || "").trim(),
      "EmergencyPhone": (datos.EmergencyPhone || "").trim()
    };

    // Aplicar las ediciones en el rango correspondiente
    Object.keys(updates).forEach(header => {
      const colIdx = userHeaders.indexOf(header);
      if (colIdx !== -1) {
        sheetUsers.getRange(foundRowIdx, colIdx + 1).setValue(updates[header]);
      }
    });

    // Guardar un log de auditoría
    try {
      const sheetLogs = ss.getSheetByName("Logs_Audit");
      if (sheetLogs) {
        const nowStr = new Date().toISOString().replace("T", " ").substring(0, 19);
        sheetLogs.appendRow([nowStr, emailClean, "UPDATE_PUBLIC_PROFILE", "Actualizó sus datos personales de forma pública vía WhatsApp."]);
      }
    } catch (e) {
      console.warn("No se pudo escribir en Logs_Audit:", e);
    }

    return { success: true, message: "Datos actualizados con éxito." };
  } catch (error) {
    console.error("Error en actualizarDatosSocioPublico:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Verifica si un DNI ya existe en la hoja de usuarios.
 */
function chequearDniSocioNuevo(dni) {
  try {
    const ss = getSpreadsheet();
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheetUsers) throw new Error("No se encontró la hoja de usuarios.");

    const dniClean = (dni || "").toString().replace(/[^0-9]/g, "").trim();
    if (!dniClean) {
      return { success: false, message: "DNI inválido." };
    }

    const data = sheetUsers.getDataRange().getValues();
    const headers = data[0];
    const dniColIndex = headers.indexOf("DNI");

    if (dniColIndex === -1) {
      throw new Error("Estructura de la hoja de usuarios inválida.");
    }

    let existe = false;
    for (let i = 1; i < data.length; i++) {
      const currentDni = (data[i][dniColIndex] || "").toString().replace(/[^0-9]/g, "").trim();
      if (currentDni === dniClean) {
        existe = true;
        break;
      }
    }

    return { success: true, existe: existe };
  } catch (error) {
    console.error("Error en chequearDniSocioNuevo:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Obtiene la lista de categorías registradas en el club para uso público.
 */
function obtenerCategoriasPublicas() {
  try {
    const ss = getSpreadsheet();
    const sheetCats = ss.getSheetByName(HOJA_CATEGORIAS);
    if (!sheetCats) throw new Error("No se encontró la hoja de categorías.");

    const data = sheetCats.getDataRange().getValues();
    const headers = data[0];
    const idColIdx = headers.indexOf("Category_ID");
    const nameColIdx = headers.indexOf("Name");

    if (idColIdx === -1 || nameColIdx === -1) {
      throw new Error("Estructura de la hoja de categorías inválida.");
    }

    const categorias = [];
    for (let i = 1; i < data.length; i++) {
      const id = (data[i][idColIdx] || "").toString().trim();
      const name = (data[i][nameColIdx] || "").toString().trim();
      if (id && name) {
        categorias.push({ id: id, name: name });
      }
    }

    return { success: true, categorias: categorias };
  } catch (error) {
    console.error("Error en obtenerCategoriasPublicas:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Registra un nuevo socio de forma pública desde el formulario de WhatsApp.
 * Autogenera usuario único y define la clave '1234' por defecto.
 */
function registrarSocioPublico(socioObj) {
  try {
    const ss = getSpreadsheet();
    const sheetUsers = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheetUsers) throw new Error("No se encontró la hoja de usuarios.");

    const email = (socioObj.Email || "").trim().toLowerCase();
    const dni = (socioObj.DNI || "").toString().replace(/[^0-9]/g, "").trim();

    if (!email) throw new Error("El correo electrónico es obligatorio.");
    if (dni.length !== 8 || isNaN(dni)) throw new Error("El DNI debe tener exactamente 8 caracteres numéricos.");

    // Verificar si ya existe email o DNI
    const data = sheetUsers.getDataRange().getValues();
    const headers = data[0];
    const emailColIdx = headers.indexOf("Email");
    const dniColIdx = headers.indexOf("DNI");
    const userColIdx = headers.indexOf("Username");

    for (let i = 1; i < data.length; i++) {
      if (emailColIdx !== -1 && (data[i][emailColIdx] || "").toString().toLowerCase().trim() === email) {
        throw new Error("Ya existe un socio registrado con el correo " + email);
      }
      if (dniColIdx !== -1 && (data[i][dniColIdx] || "").toString().replace(/[^0-9]/g, "").trim() === dni) {
        throw new Error("Ya existe un socio registrado con el DNI " + dni);
      }
    }

    // Sanitizar nombres y apellidos para generar usuario
    function sanitizar(txt) {
      return (txt || "").toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "") // remover acentos
        .replace(/[^a-z0-9]/g, "");    // alfanumérico únicamente
    }

    const firstName = socioObj.FirstName || "";
    const lastName = socioObj.LastName || "";

    const apellidoClean = sanitizar(lastName);
    const nombresArr = firstName.split(/\s+/).map(n => sanitizar(n)).filter(Boolean);

    if (nombresArr.length === 0 || !apellidoClean) {
      throw new Error("Nombre o Apellido inválidos.");
    }

    // Obtener lista de usuarios existentes en mayúsculas
    const existingUsernames = new Set();
    for (let i = 1; i < data.length; i++) {
      if (userColIdx !== -1 && data[i][userColIdx]) {
        existingUsernames.add(data[i][userColIdx].toString().trim().toUpperCase());
      }
    }

    // Algoritmo de 5 combinaciones
    const combinaciones = [];
    const p1 = (nombresArr[0][0] || "") + apellidoClean;
    combinaciones.push(p1.toUpperCase());

    if (nombresArr.length > 1) {
      const p2 = nombresArr.map(n => n[0] || "").join("") + apellidoClean;
      combinaciones.push(p2.toUpperCase());
    }

    if (nombresArr[0].length >= 2) {
      const p3 = nombresArr[0].substring(0, 2) + apellidoClean;
      combinaciones.push(p3.toUpperCase());
    }

    if (nombresArr[0].length >= 3) {
      const p4 = nombresArr[0].substring(0, 3) + apellidoClean;
      combinaciones.push(p4.toUpperCase());
    }

    let candidato = "";
    for (const comb of combinaciones) {
      if (!existingUsernames.has(comb)) {
        candidato = comb;
        break;
      }
    }

    if (!candidato) {
      let i = 1;
      const base = ((nombresArr[0][0] || "") + apellidoClean).toUpperCase();
      while (true) {
        const combNum = base + i;
        if (!existingUsernames.has(combNum)) {
          candidato = combNum;
          break;
        }
        i++;
      }
    }

    // Formatear nombre completo para la base (Apellido, Nombres)
    const formattedName = lastName.trim().split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(" ") + ", " +
                          firstName.trim().split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(" ");

    // Calcular edad
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

    // Formatear fila con los headers correctos
    const nowStr = new Date().toISOString().substring(0, 10); // YYYY-MM-DD
    const newRow = [];
    headers.forEach(header => {
      if (header === "Email") newRow.push(email);
      else if (header === "Role") newRow.push("Deportista");
      else if (header === "Name") newRow.push(formattedName);
      else if (header === "Phone") newRow.push((socioObj.Phone || "").trim());
      else if (header === "Category") newRow.push(socioObj.Category || "");
      else if (header === "DNI") newRow.push(dni);
      else if (header === "BirthDate") newRow.push(socioObj.BirthDate || "");
      else if (header === "Age") newRow.push(age.toString());
      else if (header === "JoinDate") newRow.push(nowStr);
      else if (header === "BloodType") newRow.push(socioObj.BloodType || "O+");
      else if (header === "MedicalFit") newRow.push(socioObj.MedicalFit || "Apto Físico Vigente");
      else if (header === "ObraSocial") newRow.push(socioObj.ObraSocial || "Particular");
      else if (header === "EmergencyContact") newRow.push(socioObj.EmergencyContact || "");
      else if (header === "EmergencyPhone") newRow.push(socioObj.EmergencyPhone || "");
      else if (header === "Username") newRow.push(candidato);
      else if (header === "Password") newRow.push("1234");
      else newRow.push("");
    });

    sheetUsers.appendRow(newRow);

    // Escribir en Logs_Audit
    try {
      const sheetLogs = ss.getSheetByName("Logs_Audit");
      if (sheetLogs) {
        const timeStr = new Date().toISOString().replace("T", " ").substring(0, 19);
        sheetLogs.appendRow([timeStr, email, "REGISTER_PUBLIC_SOCIO", "Se registró como nuevo socio de forma pública. Usuario autogenerado: " + candidato]);
      }
    } catch (e) {
      console.warn("No se pudo escribir en Logs_Audit:", e);
    }

    return { success: true, username: candidato, password: "1234" };
  } catch (error) {
    console.error("Error en registrarSocioPublico:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Actualiza la clave de acceso numérica de un usuario.
 */
function actualizarClaveUsuario(email, nuevaClave) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheet) throw new Error("No se encontró la hoja de usuarios.");

    const emailClean = (email || "").trim().toLowerCase();
    const claveClean = (nuevaClave || "").toString().trim();

    if (claveClean.length !== 4 || isNaN(claveClean)) {
      throw new Error("La clave debe tener exactamente 4 dígitos numéricos.");
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailColIdx = headers.indexOf("Email");
    const passColIdx = headers.indexOf("Password");

    if (emailColIdx === -1 || passColIdx === -1) {
      throw new Error("Estructura de la hoja de usuarios inválida.");
    }

    let foundRowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if ((data[i][emailColIdx] || "").toString().toLowerCase().trim() === emailClean) {
        foundRowIdx = i + 1; // +1 for headers
        break;
      }
    }

    if (foundRowIdx === -1) {
      throw new Error("No se encontró ningún usuario con el correo especificado.");
    }

    // Actualizar clave en la hoja
    sheet.getRange(foundRowIdx, passColIdx + 1).setValue(claveClean);

    return { success: true, message: "Clave actualizada con éxito." };
  } catch (error) {
    console.error("Error en actualizarClaveUsuario:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Genera un token de restablecimiento y envía un correo con el link correspondiente.
 */
function solicitarRestablecimientoClave(email, origin) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_USUARIOS);
    if (!sheet) throw new Error("No se encontró la hoja de usuarios.");

    const emailClean = (email || "").trim().toLowerCase();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailColIdx = headers.indexOf("Email");
    const nameColIdx = headers.indexOf("Name");

    if (emailColIdx === -1) {
      throw new Error("Estructura de la hoja de usuarios inválida.");
    }

    let foundRowIdx = -1;
    let userName = "Socio";
    for (let i = 1; i < data.length; i++) {
      if ((data[i][emailColIdx] || "").toString().toLowerCase().trim() === emailClean) {
        foundRowIdx = i + 1;
        if (nameColIdx !== -1) {
          userName = data[i][nameColIdx] || "Socio";
        }
        break;
      }
    }

    if (foundRowIdx === -1) {
      return { success: false, message: "El correo ingresado no pertenece a ningún socio registrado." };
    }

    // Generar un token aleatorio simple
    const token = Math.floor(100000 + Math.random() * 900000).toString();
    
    // Guardar token en el cache
    const cache = CacheService.getScriptCache();
    cache.put("reset-token-" + emailClean, token, 900); // 15 minutos

    // Construir el link
    const resetLink = origin + "?action=reset&email=" + encodeURIComponent(emailClean) + "&token=" + token;

    // Enviar el correo
    const subject = "Restablecer Clave - Haedo Futsal";
    const body = "Hola " + userName + ",\n\n" +
                 "Recibimos una solicitud para restablecer la clave numérica de tu cuenta en Haedo Futsal.\n\n" +
                 "Para ingresar tu nueva clave, hacé click en el siguiente enlace (enlace válido por 15 minutos):\n" +
                 resetLink + "\n\n" +
                 "Si no solicitaste este cambio, podés ignorar este correo de forma segura.\n\n" +
                 "Saludos,\n" +
                 "Administración - Haedo Futsal";

    MailApp.sendEmail(emailClean, subject, body);

    return { success: true, message: "Correo de restablecimiento enviado." };
  } catch (error) {
    console.error("Error en solicitarRestablecimientoClave:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Valida el token y restablece la clave del usuario.
 */
function restablecerClaveConToken(email, token, nuevaClave) {
  try {
    const emailClean = (email || "").trim().toLowerCase();
    const tokenClean = (token || "").trim();
    const claveClean = (nuevaClave || "").toString().trim();

    if (claveClean.length !== 4 || isNaN(claveClean)) {
      throw new Error("La clave debe tener exactamente 4 dígitos numéricos.");
    }

    // Verificar token en el cache
    const cache = CacheService.getScriptCache();
    const cachedToken = cache.get("reset-token-" + emailClean);

    if (!cachedToken || cachedToken !== tokenClean) {
      throw new Error("El enlace de restablecimiento es inválido o ha expirado.");
    }

    // Actualizar la clave
    const res = actualizarClaveUsuario(emailClean, claveClean);
    if (!res.success) throw new Error(res.message);

    // Borrar token del cache
    cache.remove("reset-token-" + emailClean);

    return { success: true, message: "Clave restablecida con éxito." };
  } catch (error) {
    console.error("Error en restablecerClaveConToken:", error);
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
        title: `Haedo Futsal - Cuota Mes: ${month}`,
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
 * de prueba ampliados (con Torneos, Finanzas y Partidos) para Haedo Futsal.
 */
function inicializarBaseDatos() {
  const ss = getSpreadsheet();
  
  // 1. Configurar Hoja "Admins"
  let sheetAdmins = ss.getSheetByName(HOJA_ADMINS) || ss.insertSheet(HOJA_ADMINS);
  sheetAdmins.clear();
  sheetAdmins.appendRow(["Email"]);
  const admins = [["admin@futsalhaedo.com"],["guillermopazos@gmail.com"]];
  admins.forEach(r => sheetAdmins.appendRow(r));
  
  // 2. Configurar Hoja "Categorias"
  let sheetCats = ss.getSheetByName(HOJA_CATEGORIAS) || ss.insertSheet(HOJA_CATEGORIAS);
  sheetCats.clear();
  sheetCats.appendRow(["Category_ID","Name","Coach","Monthly_Fee","Torneos"]);
  const cats = [["Fut-May+35","Futsal Mayores +35","Bocha",15000,"EDEFI"],["Fut-May+42","Futsal Mayores +42","Bocha",15000,"EDEFI"],["Fut-Fem-1ra","Futsal Femenino 1ra","Pol",20000,"FutFEM"],["Fut-Juv-5ta","Futsal Juveniles 5ta","Bocha",30000,"FUTSALA"],["Bab-2015","Baby Futbol - 2015","Marcelo",30000,"Baby EDEFI"]];
  cats.forEach(r => sheetCats.appendRow(r));
  
  // 3. Configurar Hoja "Usuarios"
  let sheetUsers = ss.getSheetByName(HOJA_USUARIOS) || ss.insertSheet(HOJA_USUARIOS);
  sheetUsers.clear();
  sheetUsers.appendRow(["Email","Role","Name","Phone","Category","DNI","BirthDate","Age","JoinDate","BloodType","MedicalFit","ObraSocial","EmergencyContact","EmergencyPhone","ParentName","ParentPhone","Notes","Photo","Username","Password"]);
  const users = [["admin@futsalhaedo.com","Admin","Administrador Haedo","1122334455","Bab-2015","35.100.200","2015-01-15","11","2024-03-01","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Haedo Roberto (Hermano/Familiar)","11-3000-4000","Haedo Alberto / Maria (Padres)","11-5000-6000","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=adminfutsalhaedocom"],["guillermopazos@gmail.com","Admin","Guillermo Pazos","1133445566","Bab-2015","36.101.201","2015-02-15","11","2024-03-02","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Pazos Roberto (Hermano/Familiar)","11-3001-4001","Pazos Alberto / Maria (Padres)","11-5001-6001","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=guillermopazosgmailcom"],["deportista@futsalhaedo.com","Deportista","Juan Perez","1199887766","Fut-May+35","37.102.202","1988-03-15","38","2024-03-03","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Perez Roberto (Hermano/Familiar)","11-3002-4002","Perez Alberto / Maria (Padres)","11-5002-6002","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=deportistafutsalhaedocom"],["socio1@futsalhaedo.com","Deportista","Sofia Messi","1120005432","Fut-May+35","38.103.203","1988-04-15","38","2024-03-04","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Messi Roberto (Hermano/Familiar)","11-3003-4003","Messi Alberto / Maria (Padres)","11-5003-6003","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio1futsalhaedocom"],["socio2@futsalhaedo.com","Deportista","Javier Lopez","1120010864","Fut-May+42","39.104.204","1980-05-15","46","2024-03-05","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Lopez Roberto (Hermano/Familiar)","11-3004-4004","Lopez Alberto / Maria (Padres)","11-5004-6004","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio2futsalhaedocom"],["socio3@futsalhaedo.com","Deportista","Lautaro Aguero","1120016296","Fut-Fem-1ra","40.105.205","2000-06-15","26","2024-03-06","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Aguero Roberto (Hermano/Familiar)","11-3005-4005","Aguero Alberto / Maria (Padres)","11-5005-6005","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio3futsalhaedocom"],["socio4@futsalhaedo.com","Deportista","Yamila Correa","1120021728","Fut-Juv-5ta","41.106.206","2008-07-15","18","2024-03-07","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Correa Roberto (Hermano/Familiar)","11-3006-4006","Correa Alberto / Maria (Padres)","11-5006-6006","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio4futsalhaedocom"],["socio5@futsalhaedo.com","Deportista","Alexis Lopez","1120027160","Bab-2015","42.107.207","2015-08-15","11","2024-03-08","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Lopez Roberto (Hermano/Familiar)","11-3007-4007","Lopez Alberto / Maria (Padres)","11-5007-6007","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio5futsalhaedocom"],["socio6@futsalhaedo.com","Deportista","Vanina Rodriguez","1120032592","Fut-May+35","43.108.208","1988-09-15","38","2024-03-09","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Rodriguez Roberto (Hermano/Familiar)","11-3008-4008","Rodriguez Alberto / Maria (Padres)","11-5008-6008","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio6futsalhaedocom"],["socio7@futsalhaedo.com","Deportista","Valentina Torres","1120038024","Fut-May+42","44.109.209","1980-01-15","46","2024-03-10","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Torres Roberto (Hermano/Familiar)","11-3009-4009","Torres Alberto / Maria (Padres)","11-5009-6009","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio7futsalhaedocom"],["socio8@futsalhaedo.com","Deportista","Joaquin Mac Allister","1120043456","Fut-Fem-1ra","45.110.210","2000-02-15","26","2024-03-11","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Allister Roberto (Hermano/Familiar)","11-3010-4010","Allister Alberto / Maria (Padres)","11-5010-6010","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio8futsalhaedocom"],["socio9@futsalhaedo.com","Deportista","Roman Francescoli","1120048888","Fut-Juv-5ta","46.111.211","2008-03-15","18","2024-03-12","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Francescoli Roberto (Hermano/Familiar)","11-3011-4011","Francescoli Alberto / Maria (Padres)","11-5011-6011","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio9futsalhaedocom"],["socio10@futsalhaedo.com","Deportista","Tomas Palermo","1120054320","Bab-2015","47.112.212","2015-04-15","11","2024-03-13","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Palermo Roberto (Hermano/Familiar)","11-3012-4012","Palermo Alberto / Maria (Padres)","11-5012-6012","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio10futsalhaedocom"],["socio11@futsalhaedo.com","Deportista","Javier Messi","1120059752","Fut-May+35","48.113.213","1988-05-15","38","2024-03-14","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Messi Roberto (Hermano/Familiar)","11-3013-4013","Messi Alberto / Maria (Padres)","11-5013-6013","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio11futsalhaedocom"],["socio12@futsalhaedo.com","Deportista","Valentina Gonzalez","1120065184","Fut-May+42","49.114.214","1980-06-15","46","2024-03-15","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Gonzalez Roberto (Hermano/Familiar)","11-3014-4014","Gonzalez Alberto / Maria (Padres)","11-5014-6014","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio12futsalhaedocom"],["socio13@futsalhaedo.com","Deportista","Lionel Perez","1120070616","Fut-Fem-1ra","35.115.215","2000-07-15","26","2024-03-16","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Perez Roberto (Hermano/Familiar)","11-3015-4015","Perez Alberto / Maria (Padres)","11-5015-6015","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio13futsalhaedocom"],["socio14@futsalhaedo.com","Deportista","Alexis Ronaldo","1120076048","Fut-Juv-5ta","36.116.216","2008-08-15","18","2024-03-17","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Ronaldo Roberto (Hermano/Familiar)","11-3016-4016","Ronaldo Alberto / Maria (Padres)","11-5016-6016","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio14futsalhaedocom"],["socio15@futsalhaedo.com","Deportista","Carlos Rodriguez","1120081480","Bab-2015","37.117.217","2015-09-15","11","2024-03-18","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Rodriguez Roberto (Hermano/Familiar)","11-3017-4017","Rodriguez Alberto / Maria (Padres)","11-5017-6017","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio15futsalhaedocom"],["socio16@futsalhaedo.com","Deportista","Mateo Rodriguez","1120086912","Fut-May+35","38.118.218","1988-01-15","38","2024-03-19","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Rodriguez Roberto (Hermano/Familiar)","11-3018-4018","Rodriguez Alberto / Maria (Padres)","11-5018-6018","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio16futsalhaedocom"],["socio17@futsalhaedo.com","Deportista","Enzo Martinez","1120092344","Fut-May+42","39.119.219","1980-02-15","46","2024-03-20","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Martinez Roberto (Hermano/Familiar)","11-3019-4019","Martinez Alberto / Maria (Padres)","11-5019-6019","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio17futsalhaedocom"],["socio18@futsalhaedo.com","Deportista","Vanina Zanetti","1120097776","Fut-Fem-1ra","40.120.220","2000-03-15","26","2024-03-21","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Zanetti Roberto (Hermano/Familiar)","11-3020-4020","Zanetti Alberto / Maria (Padres)","11-5020-6020","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio18futsalhaedocom"],["socio19@futsalhaedo.com","Deportista","Mateo Lopez","1120103208","Fut-Juv-5ta","41.121.221","2008-04-15","18","2024-03-22","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Lopez Roberto (Hermano/Familiar)","11-3021-4021","Lopez Alberto / Maria (Padres)","11-5021-6021","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio19futsalhaedocom"],["socio20@futsalhaedo.com","Deportista","Geronimo Gomez","1120108640","Bab-2015","42.122.222","2015-05-15","11","2024-03-23","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Gomez Roberto (Hermano/Familiar)","11-3022-4022","Gomez Alberto / Maria (Padres)","11-5022-6022","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio20futsalhaedocom"],["socio21@futsalhaedo.com","Deportista","Benjamin Gonzalez","1120114072","Fut-May+35","43.123.223","1988-06-15","38","2024-03-24","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Gonzalez Roberto (Hermano/Familiar)","11-3023-4023","Gonzalez Alberto / Maria (Padres)","11-5023-6023","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio21futsalhaedocom"],["socio22@futsalhaedo.com","Deportista","Yamila Ronaldo","1120119504","Fut-May+42","44.124.224","1980-07-15","46","2024-03-25","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Ronaldo Roberto (Hermano/Familiar)","11-3024-4024","Ronaldo Alberto / Maria (Padres)","11-5024-6024","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio22futsalhaedocom"],["socio23@futsalhaedo.com","Deportista","Julian Ronaldo","1120124936","Fut-Fem-1ra","45.125.225","2000-08-15","26","2024-03-01","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Ronaldo Roberto (Hermano/Familiar)","11-3025-4025","Ronaldo Alberto / Maria (Padres)","11-5025-6025","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio23futsalhaedocom"],["socio24@futsalhaedo.com","Deportista","Geronimo Sosa","1120130368","Fut-Juv-5ta","46.126.226","2008-09-15","18","2024-03-02","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Sosa Roberto (Hermano/Familiar)","11-3026-4026","Sosa Alberto / Maria (Padres)","11-5026-6026","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio24futsalhaedocom"],["socio25@futsalhaedo.com","Deportista","Tomas Gonzalez","1120135800","Bab-2015","47.127.227","2015-01-15","11","2024-03-03","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Gonzalez Roberto (Hermano/Familiar)","11-3027-4027","Gonzalez Alberto / Maria (Padres)","11-5027-6027","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio25futsalhaedocom"],["socio26@futsalhaedo.com","Deportista","Geronimo Bonsegundo","1120141232","Fut-May+35","48.128.228","1988-02-15","38","2024-03-04","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Bonsegundo Roberto (Hermano/Familiar)","11-3028-4028","Bonsegundo Alberto / Maria (Padres)","11-5028-6028","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio26futsalhaedocom"],["socio27@futsalhaedo.com","Deportista","Tomas Bonsegundo","1120146664","Fut-May+42","49.129.229","1980-03-15","46","2024-03-05","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Bonsegundo Roberto (Hermano/Familiar)","11-3029-4029","Bonsegundo Alberto / Maria (Padres)","11-5029-6029","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio27futsalhaedocom"],["socio28@futsalhaedo.com","Deportista","Franco Zanetti","1120152096","Fut-Fem-1ra","35.130.230","2000-04-15","26","2024-03-06","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Zanetti Roberto (Hermano/Familiar)","11-3030-4030","Zanetti Alberto / Maria (Padres)","11-5030-6030","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio28futsalhaedocom"],["socio29@futsalhaedo.com","Deportista","Benjamin Francescoli","1120157528","Fut-Juv-5ta","36.131.231","2008-05-15","18","2024-03-07","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Francescoli Roberto (Hermano/Familiar)","11-3031-4031","Francescoli Alberto / Maria (Padres)","11-5031-6031","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio29futsalhaedocom"],["socio30@futsalhaedo.com","Deportista","Javier Perez","1120162960","Bab-2015","37.132.232","2015-06-15","11","2024-03-08","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Perez Roberto (Hermano/Familiar)","11-3032-4032","Perez Alberto / Maria (Padres)","11-5032-6032","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio30futsalhaedocom"],["socio31@futsalhaedo.com","Deportista","Agustina Banini","1120168392","Fut-May+35","38.133.233","1988-07-15","38","2024-03-09","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Banini Roberto (Hermano/Familiar)","11-3033-4033","Banini Alberto / Maria (Padres)","11-5033-6033","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio31futsalhaedocom"],["socio32@futsalhaedo.com","Deportista","Camila Alvarez","1120173824","Fut-May+42","39.134.234","1980-08-15","46","2024-03-10","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Alvarez Roberto (Hermano/Familiar)","11-3034-4034","Alvarez Alberto / Maria (Padres)","11-5034-6034","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio32futsalhaedocom"],["socio33@futsalhaedo.com","Deportista","Mateo Banini","1120179256","Fut-Fem-1ra","40.135.235","2000-09-15","26","2024-03-11","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Banini Roberto (Hermano/Familiar)","11-3035-4035","Banini Alberto / Maria (Padres)","11-5035-6035","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio33futsalhaedocom"],["socio34@futsalhaedo.com","Deportista","Javier Francescoli","1120184688","Fut-Juv-5ta","41.136.236","2008-01-15","18","2024-03-12","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Francescoli Roberto (Hermano/Familiar)","11-3036-4036","Francescoli Alberto / Maria (Padres)","11-5036-6036","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio34futsalhaedocom"],["socio35@futsalhaedo.com","Deportista","Lucia Garnacho","1120190120","Bab-2015","42.137.237","2015-02-15","11","2024-03-13","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Garnacho Roberto (Hermano/Familiar)","11-3037-4037","Garnacho Alberto / Maria (Padres)","11-5037-6037","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio35futsalhaedocom"],["socio36@futsalhaedo.com","Deportista","Florencia Francescoli","1120195552","Fut-May+35","43.138.238","1988-03-15","38","2024-03-14","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Francescoli Roberto (Hermano/Familiar)","11-3038-4038","Francescoli Alberto / Maria (Padres)","11-5038-6038","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio36futsalhaedocom"],["socio37@futsalhaedo.com","Deportista","Roman Mac Allister","1120200984","Fut-May+42","44.139.239","1980-04-15","46","2024-03-15","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Allister Roberto (Hermano/Familiar)","11-3039-4039","Allister Alberto / Maria (Padres)","11-5039-6039","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio37futsalhaedocom"],["socio38@futsalhaedo.com","Deportista","Camila Ruiz","1120206416","Fut-Fem-1ra","45.140.240","2000-05-15","26","2024-03-16","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Ruiz Roberto (Hermano/Familiar)","11-3040-4040","Ruiz Alberto / Maria (Padres)","11-5040-6040","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio38futsalhaedocom"],["socio39@futsalhaedo.com","Deportista","Bautista Perez","1120211848","Fut-Juv-5ta","46.141.241","2008-06-15","18","2024-03-17","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Perez Roberto (Hermano/Familiar)","11-3041-4041","Perez Alberto / Maria (Padres)","11-5041-6041","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio39futsalhaedocom"],["socio40@futsalhaedo.com","Deportista","Lucia Alonso","1120217280","Bab-2015","47.142.242","2015-07-15","11","2024-03-18","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Alonso Roberto (Hermano/Familiar)","11-3042-4042","Alonso Alberto / Maria (Padres)","11-5042-6042","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio40futsalhaedocom"],["socio41@futsalhaedo.com","Deportista","Benjamin Gomez","1120222712","Fut-May+35","48.143.243","1988-08-15","38","2024-03-19","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Gomez Roberto (Hermano/Familiar)","11-3043-4043","Gomez Alberto / Maria (Padres)","11-5043-6043","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio41futsalhaedocom"],["socio42@futsalhaedo.com","Deportista","Agustina Ortega","1120228144","Fut-May+42","49.144.244","1980-09-15","46","2024-03-20","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Ortega Roberto (Hermano/Familiar)","11-3044-4044","Ortega Alberto / Maria (Padres)","11-5044-6044","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio42futsalhaedocom"],["socio43@futsalhaedo.com","Deportista","Sofia Diaz","1120233576","Fut-Fem-1ra","35.145.245","2000-01-15","26","2024-03-21","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Diaz Roberto (Hermano/Familiar)","11-3045-4045","Diaz Alberto / Maria (Padres)","11-5045-6045","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio43futsalhaedocom"],["socio44@futsalhaedo.com","Deportista","Javier Maradona","1120239008","Fut-Juv-5ta","36.146.246","2008-02-15","18","2024-03-22","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Maradona Roberto (Hermano/Familiar)","11-3046-4046","Maradona Alberto / Maria (Padres)","11-5046-6046","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio44futsalhaedocom"],["socio45@futsalhaedo.com","Deportista","Milagros Aguero","1120244440","Bab-2015","37.147.247","2015-03-15","11","2024-03-23","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Aguero Roberto (Hermano/Familiar)","11-3047-4047","Aguero Alberto / Maria (Padres)","11-5047-6047","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio45futsalhaedocom"],["socio46@futsalhaedo.com","Deportista","Enzo Gomez","1120249872","Fut-May+35","38.148.248","1988-04-15","38","2024-03-24","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Gomez Roberto (Hermano/Familiar)","11-3048-4048","Gomez Alberto / Maria (Padres)","11-5048-6048","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio46futsalhaedocom"],["socio47@futsalhaedo.com","Deportista","Yamila Romero","1120255304","Fut-May+42","39.149.249","1980-05-15","46","2024-03-25","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Romero Roberto (Hermano/Familiar)","11-3049-4049","Romero Alberto / Maria (Padres)","11-5049-6049","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio47futsalhaedocom"],["socio48@futsalhaedo.com","Deportista","Joaquin Lopez","1120260736","Fut-Fem-1ra","40.150.250","2000-06-15","26","2024-03-01","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Lopez Roberto (Hermano/Familiar)","11-3050-4050","Lopez Alberto / Maria (Padres)","11-5050-6050","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio48futsalhaedocom"],["socio49@futsalhaedo.com","Deportista","Florencia Perez","1120266168","Fut-Juv-5ta","41.151.251","2008-07-15","18","2024-03-02","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Perez Roberto (Hermano/Familiar)","11-3051-4051","Perez Alberto / Maria (Padres)","11-5051-6051","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio49futsalhaedocom"],["socio50@futsalhaedo.com","Deportista","Geronimo Bianchi","1120271600","Bab-2015","42.152.252","2015-08-15","11","2024-03-03","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Bianchi Roberto (Hermano/Familiar)","11-3052-4052","Bianchi Alberto / Maria (Padres)","11-5052-6052","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio50futsalhaedocom"],["socio51@futsalhaedo.com","Deportista","Bautista Diaz","1120277032","Fut-May+35","43.153.253","1988-09-15","38","2024-03-04","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Diaz Roberto (Hermano/Familiar)","11-3053-4053","Diaz Alberto / Maria (Padres)","11-5053-6053","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio51futsalhaedocom"],["socio52@futsalhaedo.com","Deportista","Joaquin Riquelme","1120282464","Fut-May+42","44.154.254","1980-01-15","46","2024-03-05","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Riquelme Roberto (Hermano/Familiar)","11-3054-4054","Riquelme Alberto / Maria (Padres)","11-5054-6054","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio52futsalhaedocom"],["socio53@futsalhaedo.com","Deportista","Lucia Gomez","1120287896","Fut-Fem-1ra","45.155.255","2000-02-15","26","2024-03-06","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Gomez Roberto (Hermano/Familiar)","11-3055-4055","Gomez Alberto / Maria (Padres)","11-5055-6055","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio53futsalhaedocom"],["socio54@futsalhaedo.com","Deportista","Javier Lopez","1120293328","Fut-Juv-5ta","46.156.256","2008-03-15","18","2024-03-07","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Lopez Roberto (Hermano/Familiar)","11-3056-4056","Lopez Alberto / Maria (Padres)","11-5056-6056","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio54futsalhaedocom"],["socio55@futsalhaedo.com","Deportista","Agustina Rodriguez","1120298760","Bab-2015","47.157.257","2015-04-15","11","2024-03-08","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Rodriguez Roberto (Hermano/Familiar)","11-3057-4057","Rodriguez Alberto / Maria (Padres)","11-5057-6057","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio55futsalhaedocom"],["socio56@futsalhaedo.com","Deportista","Milagros Zanetti","1120304192","Fut-May+35","48.158.258","1988-05-15","38","2024-03-09","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Zanetti Roberto (Hermano/Familiar)","11-3058-4058","Zanetti Alberto / Maria (Padres)","11-5058-6058","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio56futsalhaedocom"],["socio57@futsalhaedo.com","Deportista","Sofia Fernandez","1120309624","Fut-May+42","49.159.259","1980-06-15","46","2024-03-10","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Fernandez Roberto (Hermano/Familiar)","11-3059-4059","Fernandez Alberto / Maria (Padres)","11-5059-6059","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio57futsalhaedocom"],["socio58@futsalhaedo.com","Deportista","Vanina Martinez","1120315056","Fut-Fem-1ra","35.160.260","2000-07-15","26","2024-03-11","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Martinez Roberto (Hermano/Familiar)","11-3060-4060","Martinez Alberto / Maria (Padres)","11-5060-6060","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio58futsalhaedocom"],["socio59@futsalhaedo.com","Deportista","Agustina Aguero","1120320488","Fut-Juv-5ta","36.161.261","2008-08-15","18","2024-03-12","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Aguero Roberto (Hermano/Familiar)","11-3061-4061","Aguero Alberto / Maria (Padres)","11-5061-6061","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio59futsalhaedocom"],["socio60@futsalhaedo.com","Deportista","Lionel Sosa","1120325920","Bab-2015","37.162.262","2015-09-15","11","2024-03-13","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Sosa Roberto (Hermano/Familiar)","11-3062-4062","Sosa Alberto / Maria (Padres)","11-5062-6062","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio60futsalhaedocom"],["socio61@futsalhaedo.com","Deportista","Martin Fernandez","1120331352","Fut-May+35","38.163.263","1988-01-15","38","2024-03-14","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Fernandez Roberto (Hermano/Familiar)","11-3063-4063","Fernandez Alberto / Maria (Padres)","11-5063-6063","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio61futsalhaedocom"],["socio62@futsalhaedo.com","Deportista","Estefania Bonsegundo","1120336784","Fut-May+42","39.164.264","1980-02-15","46","2024-03-15","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Bonsegundo Roberto (Hermano/Familiar)","11-3064-4064","Bonsegundo Alberto / Maria (Padres)","11-5064-6064","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio62futsalhaedocom"],["socio63@futsalhaedo.com","Deportista","Enzo Torres","1120342216","Fut-Fem-1ra","40.165.265","2000-03-15","26","2024-03-16","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Torres Roberto (Hermano/Familiar)","11-3065-4065","Torres Alberto / Maria (Padres)","11-5065-6065","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio63futsalhaedocom"],["socio64@futsalhaedo.com","Deportista","Joaquin Palermo","1120347648","Fut-Juv-5ta","41.166.266","2008-04-15","18","2024-03-17","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Palermo Roberto (Hermano/Familiar)","11-3066-4066","Palermo Alberto / Maria (Padres)","11-5066-6066","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio64futsalhaedocom"],["socio65@futsalhaedo.com","Deportista","Agustina Ruiz","1120353080","Bab-2015","42.167.267","2015-05-15","11","2024-03-18","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Ruiz Roberto (Hermano/Familiar)","11-3067-4067","Ruiz Alberto / Maria (Padres)","11-5067-6067","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio65futsalhaedocom"],["socio66@futsalhaedo.com","Deportista","Vanina Romero","1120358512","Fut-May+35","43.168.268","1988-06-15","38","2024-03-19","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Romero Roberto (Hermano/Familiar)","11-3068-4068","Romero Alberto / Maria (Padres)","11-5068-6068","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio66futsalhaedocom"],["socio67@futsalhaedo.com","Deportista","Tomas Ortega","1120363944","Fut-May+42","44.169.269","1980-07-15","46","2024-03-20","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Ortega Roberto (Hermano/Familiar)","11-3069-4069","Ortega Alberto / Maria (Padres)","11-5069-6069","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio67futsalhaedocom"],["socio68@futsalhaedo.com","Deportista","Bautista Rodriguez","1120369376","Fut-Fem-1ra","45.170.270","2000-08-15","26","2024-03-21","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Rodriguez Roberto (Hermano/Familiar)","11-3070-4070","Rodriguez Alberto / Maria (Padres)","11-5070-6070","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio68futsalhaedocom"],["socio69@futsalhaedo.com","Deportista","Bautista Martinez","1120374808","Fut-Juv-5ta","46.171.271","2008-09-15","18","2024-03-22","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Martinez Roberto (Hermano/Familiar)","11-3071-4071","Martinez Alberto / Maria (Padres)","11-5071-6071","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio69futsalhaedocom"],["socio70@futsalhaedo.com","Deportista","Franco Ortega","1120380240","Bab-2015","47.172.272","2015-01-15","11","2024-03-23","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Ortega Roberto (Hermano/Familiar)","11-3072-4072","Ortega Alberto / Maria (Padres)","11-5072-6072","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio70futsalhaedocom"],["socio71@futsalhaedo.com","Deportista","Milagros Diaz","1120385672","Fut-May+35","48.173.273","1988-02-15","38","2024-03-24","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Diaz Roberto (Hermano/Familiar)","11-3073-4073","Diaz Alberto / Maria (Padres)","11-5073-6073","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio71futsalhaedocom"],["socio72@futsalhaedo.com","Deportista","Nicolas Mac Allister","1120391104","Fut-May+42","49.174.274","1980-03-15","46","2024-03-25","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Allister Roberto (Hermano/Familiar)","11-3074-4074","Allister Alberto / Maria (Padres)","11-5074-6074","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio72futsalhaedocom"],["socio73@futsalhaedo.com","Deportista","Ariel Garnacho","1120396536","Fut-Fem-1ra","35.175.275","2000-04-15","26","2024-03-01","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Garnacho Roberto (Hermano/Familiar)","11-3075-4075","Garnacho Alberto / Maria (Padres)","11-5075-6075","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio73futsalhaedocom"],["socio74@futsalhaedo.com","Deportista","Lautaro Palermo","1120401968","Fut-Juv-5ta","36.176.276","2008-05-15","18","2024-03-02","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Palermo Roberto (Hermano/Familiar)","11-3076-4076","Palermo Alberto / Maria (Padres)","11-5076-6076","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio74futsalhaedocom"],["socio75@futsalhaedo.com","Deportista","Nicolas Torres","1120407400","Bab-2015","37.177.277","2015-06-15","11","2024-03-03","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Torres Roberto (Hermano/Familiar)","11-3077-4077","Torres Alberto / Maria (Padres)","11-5077-6077","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio75futsalhaedocom"],["socio76@futsalhaedo.com","Deportista","Julian Francescoli","1120412832","Fut-May+35","38.178.278","1988-07-15","38","2024-03-04","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Francescoli Roberto (Hermano/Familiar)","11-3078-4078","Francescoli Alberto / Maria (Padres)","11-5078-6078","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio76futsalhaedocom"],["socio77@futsalhaedo.com","Deportista","Lionel Gomez","1120418264","Fut-May+42","39.179.279","1980-08-15","46","2024-03-05","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Gomez Roberto (Hermano/Familiar)","11-3079-4079","Gomez Alberto / Maria (Padres)","11-5079-6079","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio77futsalhaedocom"],["socio78@futsalhaedo.com","Deportista","Lautaro Fernandez","1120423696","Fut-Fem-1ra","40.180.280","2000-09-15","26","2024-03-06","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Fernandez Roberto (Hermano/Familiar)","11-3080-4080","Fernandez Alberto / Maria (Padres)","11-5080-6080","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio78futsalhaedocom"],["socio79@futsalhaedo.com","Deportista","Geronimo Banini","1120429128","Fut-Juv-5ta","41.181.281","2008-01-15","18","2024-03-07","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Banini Roberto (Hermano/Familiar)","11-3081-4081","Banini Alberto / Maria (Padres)","11-5081-6081","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio79futsalhaedocom"],["socio80@futsalhaedo.com","Deportista","Florencia Fernandez","1120434560","Bab-2015","42.182.282","2015-02-15","11","2024-03-08","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Fernandez Roberto (Hermano/Familiar)","11-3082-4082","Fernandez Alberto / Maria (Padres)","11-5082-6082","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio80futsalhaedocom"],["socio81@futsalhaedo.com","Deportista","Enzo Bonsegundo","1120439992","Fut-May+35","43.183.283","1988-03-15","38","2024-03-09","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Bonsegundo Roberto (Hermano/Familiar)","11-3083-4083","Bonsegundo Alberto / Maria (Padres)","11-5083-6083","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio81futsalhaedocom"],["socio82@futsalhaedo.com","Deportista","Ariel Rodriguez","1120445424","Fut-May+42","44.184.284","1980-04-15","46","2024-03-10","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Rodriguez Roberto (Hermano/Familiar)","11-3084-4084","Rodriguez Alberto / Maria (Padres)","11-5084-6084","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio82futsalhaedocom"],["socio83@futsalhaedo.com","Deportista","Florencia Bonsegundo","1120450856","Fut-Fem-1ra","45.185.285","2000-05-15","26","2024-03-11","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Bonsegundo Roberto (Hermano/Familiar)","11-3085-4085","Bonsegundo Alberto / Maria (Padres)","11-5085-6085","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio83futsalhaedocom"],["socio84@futsalhaedo.com","Deportista","Lautaro Alvarez","1120456288","Fut-Juv-5ta","46.186.286","2008-06-15","18","2024-03-12","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Alvarez Roberto (Hermano/Familiar)","11-3086-4086","Alvarez Alberto / Maria (Padres)","11-5086-6086","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio84futsalhaedocom"],["socio85@futsalhaedo.com","Deportista","Thiago Gonzalez","1120461720","Bab-2015","47.187.287","2015-07-15","11","2024-03-13","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Gonzalez Roberto (Hermano/Familiar)","11-3087-4087","Gonzalez Alberto / Maria (Padres)","11-5087-6087","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio85futsalhaedocom"],["socio86@futsalhaedo.com","Deportista","Julian Alonso","1120467152","Fut-May+35","48.188.288","1988-08-15","38","2024-03-14","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Alonso Roberto (Hermano/Familiar)","11-3088-4088","Alonso Alberto / Maria (Padres)","11-5088-6088","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio86futsalhaedocom"],["socio87@futsalhaedo.com","Deportista","Mateo Garnacho","1120472584","Fut-May+42","49.189.289","1980-09-15","46","2024-03-15","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Garnacho Roberto (Hermano/Familiar)","11-3089-4089","Garnacho Alberto / Maria (Padres)","11-5089-6089","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio87futsalhaedocom"],["socio88@futsalhaedo.com","Deportista","Agustina Bonsegundo","1120478016","Fut-Fem-1ra","35.190.290","2000-01-15","26","2024-03-16","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Bonsegundo Roberto (Hermano/Familiar)","11-3090-4090","Bonsegundo Alberto / Maria (Padres)","11-5090-6090","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio88futsalhaedocom"],["socio89@futsalhaedo.com","Deportista","Benjamin Rodriguez","1120483448","Fut-Juv-5ta","36.191.291","2008-02-15","18","2024-03-17","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Rodriguez Roberto (Hermano/Familiar)","11-3091-4091","Rodriguez Alberto / Maria (Padres)","11-5091-6091","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio89futsalhaedocom"],["socio90@futsalhaedo.com","Deportista","Vanina Alvarez","1120488880","Bab-2015","37.192.292","2015-03-15","11","2024-03-18","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Alvarez Roberto (Hermano/Familiar)","11-3092-4092","Alvarez Alberto / Maria (Padres)","11-5092-6092","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio90futsalhaedocom"],["socio91@futsalhaedo.com","Deportista","Lucia Riquelme","1120494312","Fut-May+35","38.193.293","1988-04-15","38","2024-03-19","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Riquelme Roberto (Hermano/Familiar)","11-3093-4093","Riquelme Alberto / Maria (Padres)","11-5093-6093","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio91futsalhaedocom"],["socio92@futsalhaedo.com","Deportista","Diego Aguero","1120499744","Fut-May+42","39.194.294","1980-05-15","46","2024-03-20","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Aguero Roberto (Hermano/Familiar)","11-3094-4094","Aguero Alberto / Maria (Padres)","11-5094-6094","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio92futsalhaedocom"],["socio93@futsalhaedo.com","Deportista","Franco Romero","1120505176","Fut-Fem-1ra","40.195.295","2000-06-15","26","2024-03-21","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Romero Roberto (Hermano/Familiar)","11-3095-4095","Romero Alberto / Maria (Padres)","11-5095-6095","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio93futsalhaedocom"],["socio94@futsalhaedo.com","Deportista","Bautista Francescoli","1120510608","Fut-Juv-5ta","41.196.296","2008-07-15","18","2024-03-22","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Francescoli Roberto (Hermano/Familiar)","11-3096-4096","Francescoli Alberto / Maria (Padres)","11-5096-6096","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio94futsalhaedocom"],["socio95@futsalhaedo.com","Deportista","Mateo Zanetti","1120516040","Bab-2015","42.197.297","2015-08-15","11","2024-03-23","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Zanetti Roberto (Hermano/Familiar)","11-3097-4097","Zanetti Alberto / Maria (Padres)","11-5097-6097","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio95futsalhaedocom"],["socio96@futsalhaedo.com","Deportista","Sofia Ruiz","1120521472","Fut-May+35","43.198.298","1988-09-15","38","2024-03-24","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","Ruiz Roberto (Hermano/Familiar)","11-3098-4098","Ruiz Alberto / Maria (Padres)","11-5098-6098","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio96futsalhaedocom"],["socio97@futsalhaedo.com","Deportista","Ariel Lopez","1120526904","Fut-May+42","44.199.299","1980-01-15","46","2024-03-25","O-","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Medifé","Lopez Roberto (Hermano/Familiar)","11-3099-4099","Lopez Alberto / Maria (Padres)","11-5099-6099","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio97futsalhaedocom"],["socio98@futsalhaedo.com","Deportista","Camila Francescoli","1120532336","Fut-Fem-1ra","45.200.300","2000-02-15","26","2024-03-01","O+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","OSDE 310","Francescoli Roberto (Hermano/Familiar)","11-3100-4100","Francescoli Alberto / Maria (Padres)","11-5100-6100","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio98futsalhaedocom"],["socio99@futsalhaedo.com","Deportista","Sofia Gonzalez","1120537768","Fut-Juv-5ta","46.201.301","2008-03-15","18","2024-03-02","A+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Swiss Medical","Gonzalez Roberto (Hermano/Familiar)","11-3101-4101","Gonzalez Alberto / Maria (Padres)","11-5101-6101","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio99futsalhaedocom"],["socio100@futsalhaedo.com","Deportista","Diego Rodriguez","1120543200","Bab-2015","47.202.302","2015-04-15","11","2024-03-03","B+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","Galeno","Rodriguez Roberto (Hermano/Familiar)","11-3102-4102","Rodriguez Alberto / Maria (Padres)","11-5102-6102","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=socio100futsalhaedocom"],["coach@futsalhaedo.com","Coach","DT Marcelo Gallardo (Coach)","1155667788","Fut-May+35","48.203.303","1988-05-15","38","2024-03-04","AB+","Apto Físico Vigente (Vence: 2026-12-31 · Dr. M. Giménez)","IOMA","(Coach) Roberto (Hermano/Familiar)","11-3103-4103","(Coach) Alberto / Maria (Padres)","11-5103-6103","Ficha deportiva completa y verificada por administración.","https://i.pravatar.cc/150?u=coachfutsalhaedocom"],["carlosperez@gmail.com","Deportista","Perez, Carlos Alberto","1144556677","Fut-May+35","35.123.456","","","2026-06-28","O+","Apto Físico Vigente","Particular","","","","","",""],["marielapmolina@gmail.com","Deportista","Molina, Mariela","1199887766","Fut-May+35","38123456","1992-05-14","34","2026-06-28","O+","Apto Físico Vigente","OSDE","Contacto Emergencia","1122334455","","","Socio para pruebas de cobro",""]];
  users.forEach(r => {
    const username = r[0].split('@')[0];
    const password = "1234";
    sheetUsers.appendRow([...r, username, password]);
  });
  
  // 4. Configurar Hoja "Pagos"
  let sheetPagos = ss.getSheetByName(HOJA_PAGOS) || ss.insertSheet(HOJA_PAGOS);
  sheetPagos.clear();
  sheetPagos.appendRow(["Payment_ID","Email","Month","Amount","Status","MP_Link","Collected_By","Collected_At"]);
  const pagos = [["PAG-0001","deportista@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0001","",""],["PAG-0002","deportista@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0002","",""],["PAG-0003","deportista@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0003","",""],["PAG-0004","deportista@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0004","",""],["PAG-0005","deportista@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0005","",""],["PAG-0006","deportista@futsalhaedo.com","2026-06",15000,"Pendiente","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_for_PAG-0006","",""],["PAG-0007","socio1@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0007","",""],["PAG-0008","socio1@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0008","",""],["PAG-0009","socio1@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0009","",""],["PAG-0010","socio1@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0010","",""],["PAG-0011","socio1@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0011","",""],["PAG-0012","socio1@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0012","",""],["PAG-0013","socio2@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0013","",""],["PAG-0014","socio2@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0014","",""],["PAG-0015","socio2@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0015","",""],["PAG-0016","socio2@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0016","",""],["PAG-0017","socio2@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0017","",""],["PAG-0018","socio2@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0018","",""],["PAG-0019","socio3@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0019","",""],["PAG-0020","socio3@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0020","",""],["PAG-0021","socio3@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0021","",""],["PAG-0022","socio3@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0022","",""],["PAG-0023","socio3@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0023","",""],["PAG-0024","socio3@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0024","",""],["PAG-0025","socio4@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0025","",""],["PAG-0026","socio4@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0026","",""],["PAG-0027","socio4@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0027","",""],["PAG-0028","socio4@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0028","",""],["PAG-0029","socio4@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0029","",""],["PAG-0030","socio4@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0030","",""],["PAG-0031","socio5@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0031","",""],["PAG-0032","socio5@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0032","",""],["PAG-0033","socio5@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0033","",""],["PAG-0034","socio5@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0034","",""],["PAG-0035","socio5@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0035","",""],["PAG-0036","socio5@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0036","",""],["PAG-0037","socio6@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0037","",""],["PAG-0038","socio6@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0038","",""],["PAG-0039","socio6@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0039","",""],["PAG-0040","socio6@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0040","",""],["PAG-0041","socio6@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0041","",""],["PAG-0042","socio6@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0042","",""],["PAG-0043","socio7@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0043","",""],["PAG-0044","socio7@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0044","",""],["PAG-0045","socio7@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0045","",""],["PAG-0046","socio7@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0046","",""],["PAG-0047","socio7@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0047","",""],["PAG-0048","socio7@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0048","",""],["PAG-0049","socio8@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0049","",""],["PAG-0050","socio8@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0050","",""],["PAG-0051","socio8@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0051","",""],["PAG-0052","socio8@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0052","",""],["PAG-0053","socio8@futsalhaedo.com","2026-05",20000,"Pagado","","Coach","2026-06-27 21:50"],["PAG-0054","socio8@futsalhaedo.com","2026-06",20000,"Pagado","","Coach","2026-06-27 22:15"],["PAG-0055","socio9@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0055","",""],["PAG-0056","socio9@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0056","",""],["PAG-0057","socio9@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0057","",""],["PAG-0058","socio9@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0058","",""],["PAG-0059","socio9@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0059","",""],["PAG-0060","socio9@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0060","",""],["PAG-0061","socio10@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0061","",""],["PAG-0062","socio10@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0062","",""],["PAG-0063","socio10@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0063","",""],["PAG-0064","socio10@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0064","",""],["PAG-0065","socio10@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0065","",""],["PAG-0066","socio10@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0066","",""],["PAG-0067","socio11@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0067","",""],["PAG-0068","socio11@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0068","",""],["PAG-0069","socio11@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0069","",""],["PAG-0070","socio11@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0070","",""],["PAG-0071","socio11@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0071","",""],["PAG-0072","socio11@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0072","",""],["PAG-0073","socio12@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0073","",""],["PAG-0074","socio12@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0074","",""],["PAG-0075","socio12@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0075","",""],["PAG-0076","socio12@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0076","",""],["PAG-0077","socio12@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0077","",""],["PAG-0078","socio12@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0078","",""],["PAG-0079","socio13@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0079","",""],["PAG-0080","socio13@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0080","",""],["PAG-0081","socio13@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0081","",""],["PAG-0082","socio13@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0082","",""],["PAG-0083","socio13@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0083","",""],["PAG-0084","socio13@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0084","",""],["PAG-0085","socio14@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0085","",""],["PAG-0086","socio14@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0086","",""],["PAG-0087","socio14@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0087","",""],["PAG-0088","socio14@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0088","",""],["PAG-0089","socio14@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0089","",""],["PAG-0090","socio14@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0090","",""],["PAG-0091","socio15@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0091","",""],["PAG-0092","socio15@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0092","",""],["PAG-0093","socio15@futsalhaedo.com","2026-03",30000,"Pagado","","Coach","2026-06-27 21:45"],["PAG-0094","socio15@futsalhaedo.com","2026-04",30000,"Pendiente","","",""],["PAG-0095","socio15@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0096","socio15@futsalhaedo.com","2026-06",30000,"Pagado","","Admin","2026-06-27 22:16"],["PAG-0097","socio16@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0097","",""],["PAG-0098","socio16@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0098","",""],["PAG-0099","socio16@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0099","",""],["PAG-0100","socio16@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0100","",""],["PAG-0101","socio16@futsalhaedo.com","2026-05",15000,"Pagado","","",""],["PAG-0102","socio16@futsalhaedo.com","2026-06",15000,"Pendiente","","",""],["PAG-0103","socio17@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0103","",""],["PAG-0104","socio17@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0104","",""],["PAG-0105","socio17@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0105","",""],["PAG-0106","socio17@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0106","",""],["PAG-0107","socio17@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0107","",""],["PAG-0108","socio17@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0108","",""],["PAG-0109","socio18@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0109","",""],["PAG-0110","socio18@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0110","",""],["PAG-0111","socio18@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0111","",""],["PAG-0112","socio18@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0112","",""],["PAG-0113","socio18@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0113","",""],["PAG-0114","socio18@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0114","",""],["PAG-0115","socio19@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0115","",""],["PAG-0116","socio19@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0116","",""],["PAG-0117","socio19@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0117","",""],["PAG-0118","socio19@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0118","",""],["PAG-0119","socio19@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0119","",""],["PAG-0120","socio19@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0120","",""],["PAG-0121","socio20@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0121","",""],["PAG-0122","socio20@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0122","",""],["PAG-0123","socio20@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0123","",""],["PAG-0124","socio20@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0124","",""],["PAG-0125","socio20@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0125","",""],["PAG-0126","socio20@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0126","",""],["PAG-0127","socio21@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0127","",""],["PAG-0128","socio21@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0128","",""],["PAG-0129","socio21@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0129","",""],["PAG-0130","socio21@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0130","",""],["PAG-0131","socio21@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0131","",""],["PAG-0132","socio21@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0132","",""],["PAG-0133","socio22@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0133","",""],["PAG-0134","socio22@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0134","",""],["PAG-0135","socio22@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0135","",""],["PAG-0136","socio22@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0136","",""],["PAG-0137","socio22@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0137","",""],["PAG-0138","socio22@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0138","",""],["PAG-0139","socio23@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0139","",""],["PAG-0140","socio23@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0140","",""],["PAG-0141","socio23@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0141","",""],["PAG-0142","socio23@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0142","",""],["PAG-0143","socio23@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0143","",""],["PAG-0144","socio23@futsalhaedo.com","2026-06",20000,"Pendiente","","",""],["PAG-0145","socio24@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0145","",""],["PAG-0146","socio24@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0146","",""],["PAG-0147","socio24@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0147","",""],["PAG-0148","socio24@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0148","",""],["PAG-0149","socio24@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0150","socio24@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0151","socio25@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0151","",""],["PAG-0152","socio25@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0152","",""],["PAG-0153","socio25@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0153","",""],["PAG-0154","socio25@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0154","",""],["PAG-0155","socio25@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0155","",""],["PAG-0156","socio25@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0156","",""],["PAG-0157","socio26@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0157","",""],["PAG-0158","socio26@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0158","",""],["PAG-0159","socio26@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0159","",""],["PAG-0160","socio26@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0160","",""],["PAG-0161","socio26@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0161","",""],["PAG-0162","socio26@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0162","",""],["PAG-0163","socio27@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0163","",""],["PAG-0164","socio27@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0164","",""],["PAG-0165","socio27@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0165","",""],["PAG-0166","socio27@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0166","",""],["PAG-0167","socio27@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0167","",""],["PAG-0168","socio27@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0168","",""],["PAG-0169","socio28@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0169","",""],["PAG-0170","socio28@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0170","",""],["PAG-0171","socio28@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0171","",""],["PAG-0172","socio28@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0172","",""],["PAG-0173","socio28@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0173","",""],["PAG-0174","socio28@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0174","",""],["PAG-0175","socio29@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0175","",""],["PAG-0176","socio29@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0176","",""],["PAG-0177","socio29@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0177","",""],["PAG-0178","socio29@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0178","",""],["PAG-0179","socio29@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0179","",""],["PAG-0180","socio29@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0180","",""],["PAG-0181","socio30@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0181","",""],["PAG-0182","socio30@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0182","",""],["PAG-0183","socio30@futsalhaedo.com","2026-03",30000,"Pendiente","","",""],["PAG-0184","socio30@futsalhaedo.com","2026-04",30000,"Pendiente","","",""],["PAG-0185","socio30@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0186","socio30@futsalhaedo.com","2026-06",30000,"Pagado","","Coach","2026-06-27 22:16"],["PAG-0187","socio31@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0187","",""],["PAG-0188","socio31@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0188","",""],["PAG-0189","socio31@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0189","",""],["PAG-0190","socio31@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0190","",""],["PAG-0191","socio31@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0191","",""],["PAG-0192","socio31@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0192","",""],["PAG-0193","socio32@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0193","",""],["PAG-0194","socio32@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0194","",""],["PAG-0195","socio32@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0195","",""],["PAG-0196","socio32@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0196","",""],["PAG-0197","socio32@futsalhaedo.com","2026-05",15000,"Pendiente","","",""],["PAG-0198","socio32@futsalhaedo.com","2026-06",15000,"Pendiente","","",""],["PAG-0199","socio33@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0199","",""],["PAG-0200","socio33@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0200","",""],["PAG-0201","socio33@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0201","",""],["PAG-0202","socio33@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0202","",""],["PAG-0203","socio33@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0203","",""],["PAG-0204","socio33@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0204","",""],["PAG-0205","socio34@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0205","",""],["PAG-0206","socio34@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0206","",""],["PAG-0207","socio34@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0207","",""],["PAG-0208","socio34@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0208","",""],["PAG-0209","socio34@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0209","",""],["PAG-0210","socio34@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0210","",""],["PAG-0211","socio35@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0211","",""],["PAG-0212","socio35@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0212","",""],["PAG-0213","socio35@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0213","",""],["PAG-0214","socio35@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0214","",""],["PAG-0215","socio35@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0215","",""],["PAG-0216","socio35@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0216","",""],["PAG-0217","socio36@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0217","",""],["PAG-0218","socio36@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0218","",""],["PAG-0219","socio36@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0219","",""],["PAG-0220","socio36@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0220","",""],["PAG-0221","socio36@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0221","",""],["PAG-0222","socio36@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0222","",""],["PAG-0223","socio37@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0223","",""],["PAG-0224","socio37@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0224","",""],["PAG-0225","socio37@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0225","",""],["PAG-0226","socio37@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0226","",""],["PAG-0227","socio37@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0227","",""],["PAG-0228","socio37@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0228","",""],["PAG-0229","socio38@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0229","",""],["PAG-0230","socio38@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0230","",""],["PAG-0231","socio38@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0231","",""],["PAG-0232","socio38@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0232","",""],["PAG-0233","socio38@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0233","",""],["PAG-0234","socio38@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0234","",""],["PAG-0235","socio39@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0235","",""],["PAG-0236","socio39@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0236","",""],["PAG-0237","socio39@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0237","",""],["PAG-0238","socio39@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0238","",""],["PAG-0239","socio39@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0239","",""],["PAG-0240","socio39@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0240","",""],["PAG-0241","socio40@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0241","",""],["PAG-0242","socio40@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0242","",""],["PAG-0243","socio40@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0243","",""],["PAG-0244","socio40@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0244","",""],["PAG-0245","socio40@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0246","socio40@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0247","socio41@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0247","",""],["PAG-0248","socio41@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0248","",""],["PAG-0249","socio41@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0249","",""],["PAG-0250","socio41@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0250","",""],["PAG-0251","socio41@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0251","",""],["PAG-0252","socio41@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0252","",""],["PAG-0253","socio42@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0253","",""],["PAG-0254","socio42@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0254","",""],["PAG-0255","socio42@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0255","",""],["PAG-0256","socio42@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0256","",""],["PAG-0257","socio42@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0257","",""],["PAG-0258","socio42@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0258","",""],["PAG-0259","socio43@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0259","",""],["PAG-0260","socio43@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0260","",""],["PAG-0261","socio43@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0261","",""],["PAG-0262","socio43@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0262","",""],["PAG-0263","socio43@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0263","",""],["PAG-0264","socio43@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0264","",""],["PAG-0265","socio44@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0265","",""],["PAG-0266","socio44@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0266","",""],["PAG-0267","socio44@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0267","",""],["PAG-0268","socio44@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0268","",""],["PAG-0269","socio44@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0269","",""],["PAG-0270","socio44@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0270","",""],["PAG-0271","socio45@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0271","",""],["PAG-0272","socio45@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0272","",""],["PAG-0273","socio45@futsalhaedo.com","2026-03",30000,"Pendiente","","",""],["PAG-0274","socio45@futsalhaedo.com","2026-04",30000,"Pendiente","","",""],["PAG-0275","socio45@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0276","socio45@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0277","socio46@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0277","",""],["PAG-0278","socio46@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0278","",""],["PAG-0279","socio46@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0279","",""],["PAG-0280","socio46@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0280","",""],["PAG-0281","socio46@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0281","",""],["PAG-0282","socio46@futsalhaedo.com","2026-06",15000,"Pendiente","","",""],["PAG-0283","socio47@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0283","",""],["PAG-0284","socio47@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0284","",""],["PAG-0285","socio47@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0285","",""],["PAG-0286","socio47@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0286","",""],["PAG-0287","socio47@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0287","",""],["PAG-0288","socio47@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0288","",""],["PAG-0289","socio48@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0289","",""],["PAG-0290","socio48@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0290","",""],["PAG-0291","socio48@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0291","",""],["PAG-0292","socio48@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0292","",""],["PAG-0293","socio48@futsalhaedo.com","2026-05",20000,"Pendiente","","",""],["PAG-0294","socio48@futsalhaedo.com","2026-06",20000,"Pendiente","","",""],["PAG-0295","socio49@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0295","",""],["PAG-0296","socio49@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0296","",""],["PAG-0297","socio49@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0297","",""],["PAG-0298","socio49@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0298","",""],["PAG-0299","socio49@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0299","",""],["PAG-0300","socio49@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0300","",""],["PAG-0301","socio50@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0301","",""],["PAG-0302","socio50@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0302","",""],["PAG-0303","socio50@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0303","",""],["PAG-0304","socio50@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0304","",""],["PAG-0305","socio50@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0305","",""],["PAG-0306","socio50@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0306","",""],["PAG-0307","socio51@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0307","",""],["PAG-0308","socio51@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0308","",""],["PAG-0309","socio51@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0309","",""],["PAG-0310","socio51@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0310","",""],["PAG-0311","socio51@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0311","",""],["PAG-0312","socio51@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0312","",""],["PAG-0313","socio52@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0313","",""],["PAG-0314","socio52@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0314","",""],["PAG-0315","socio52@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0315","",""],["PAG-0316","socio52@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0316","",""],["PAG-0317","socio52@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0317","",""],["PAG-0318","socio52@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0318","",""],["PAG-0319","socio53@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0319","",""],["PAG-0320","socio53@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0320","",""],["PAG-0321","socio53@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0321","",""],["PAG-0322","socio53@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0322","",""],["PAG-0323","socio53@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0323","",""],["PAG-0324","socio53@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0324","",""],["PAG-0325","socio54@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0325","",""],["PAG-0326","socio54@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0326","",""],["PAG-0327","socio54@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0327","",""],["PAG-0328","socio54@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0328","",""],["PAG-0329","socio54@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0329","",""],["PAG-0330","socio54@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0330","",""],["PAG-0331","socio55@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0331","",""],["PAG-0332","socio55@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0332","",""],["PAG-0333","socio55@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0333","",""],["PAG-0334","socio55@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0334","",""],["PAG-0335","socio55@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0335","",""],["PAG-0336","socio55@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0336","",""],["PAG-0337","socio56@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0337","",""],["PAG-0338","socio56@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0338","",""],["PAG-0339","socio56@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0339","",""],["PAG-0340","socio56@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0340","",""],["PAG-0341","socio56@futsalhaedo.com","2026-05",15000,"Pagado","","Admin","2026-06-27 21:52"],["PAG-0342","socio56@futsalhaedo.com","2026-06",15000,"Pendiente","","",""],["PAG-0343","socio57@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0343","",""],["PAG-0344","socio57@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0344","",""],["PAG-0345","socio57@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0345","",""],["PAG-0346","socio57@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0346","",""],["PAG-0347","socio57@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0347","",""],["PAG-0348","socio57@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0348","",""],["PAG-0349","socio58@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0349","",""],["PAG-0350","socio58@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0350","",""],["PAG-0351","socio58@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0351","",""],["PAG-0352","socio58@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0352","",""],["PAG-0353","socio58@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0353","",""],["PAG-0354","socio58@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0354","",""],["PAG-0355","socio59@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0355","",""],["PAG-0356","socio59@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0356","",""],["PAG-0357","socio59@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0357","",""],["PAG-0358","socio59@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0358","",""],["PAG-0359","socio59@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0359","",""],["PAG-0360","socio59@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0360","",""],["PAG-0361","socio60@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0361","",""],["PAG-0362","socio60@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0362","",""],["PAG-0363","socio60@futsalhaedo.com","2026-03",30000,"Pendiente","","",""],["PAG-0364","socio60@futsalhaedo.com","2026-04",30000,"Pendiente","","",""],["PAG-0365","socio60@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0366","socio60@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0367","socio61@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0367","",""],["PAG-0368","socio61@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0368","",""],["PAG-0369","socio61@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0369","",""],["PAG-0370","socio61@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0370","",""],["PAG-0371","socio61@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0371","",""],["PAG-0372","socio61@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0372","",""],["PAG-0373","socio62@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0373","",""],["PAG-0374","socio62@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0374","",""],["PAG-0375","socio62@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0375","",""],["PAG-0376","socio62@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0376","",""],["PAG-0377","socio62@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0377","",""],["PAG-0378","socio62@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0378","",""],["PAG-0379","socio63@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0379","",""],["PAG-0380","socio63@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0380","",""],["PAG-0381","socio63@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0381","",""],["PAG-0382","socio63@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0382","",""],["PAG-0383","socio63@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0383","",""],["PAG-0384","socio63@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0384","",""],["PAG-0385","socio64@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0385","",""],["PAG-0386","socio64@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0386","",""],["PAG-0387","socio64@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0387","",""],["PAG-0388","socio64@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0388","",""],["PAG-0389","socio64@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0390","socio64@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0391","socio65@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0391","",""],["PAG-0392","socio65@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0392","",""],["PAG-0393","socio65@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0393","",""],["PAG-0394","socio65@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0394","",""],["PAG-0395","socio65@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0395","",""],["PAG-0396","socio65@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0396","",""],["PAG-0397","socio66@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0397","",""],["PAG-0398","socio66@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0398","",""],["PAG-0399","socio66@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0399","",""],["PAG-0400","socio66@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0400","",""],["PAG-0401","socio66@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0401","",""],["PAG-0402","socio66@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0402","",""],["PAG-0403","socio67@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0403","",""],["PAG-0404","socio67@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0404","",""],["PAG-0405","socio67@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0405","",""],["PAG-0406","socio67@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0406","",""],["PAG-0407","socio67@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0407","",""],["PAG-0408","socio67@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0408","",""],["PAG-0409","socio68@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0409","",""],["PAG-0410","socio68@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0410","",""],["PAG-0411","socio68@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0411","",""],["PAG-0412","socio68@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0412","",""],["PAG-0413","socio68@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0413","",""],["PAG-0414","socio68@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0414","",""],["PAG-0415","socio69@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0415","",""],["PAG-0416","socio69@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0416","",""],["PAG-0417","socio69@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0417","",""],["PAG-0418","socio69@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0418","",""],["PAG-0419","socio69@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0419","",""],["PAG-0420","socio69@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0421","socio70@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0421","",""],["PAG-0422","socio70@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0422","",""],["PAG-0423","socio70@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0423","",""],["PAG-0424","socio70@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0424","",""],["PAG-0425","socio70@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0425","",""],["PAG-0426","socio70@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0426","",""],["PAG-0427","socio71@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0427","",""],["PAG-0428","socio71@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0428","",""],["PAG-0429","socio71@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0429","",""],["PAG-0430","socio71@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0430","",""],["PAG-0431","socio71@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0431","",""],["PAG-0432","socio71@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0432","",""],["PAG-0433","socio72@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0433","",""],["PAG-0434","socio72@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0434","",""],["PAG-0435","socio72@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0435","",""],["PAG-0436","socio72@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0436","",""],["PAG-0437","socio72@futsalhaedo.com","2026-05",15000,"Pendiente","","",""],["PAG-0438","socio72@futsalhaedo.com","2026-06",15000,"Pendiente","","",""],["PAG-0439","socio73@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0439","",""],["PAG-0440","socio73@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0440","",""],["PAG-0441","socio73@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0441","",""],["PAG-0442","socio73@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0442","",""],["PAG-0443","socio73@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0443","",""],["PAG-0444","socio73@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0444","",""],["PAG-0445","socio74@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0445","",""],["PAG-0446","socio74@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0446","",""],["PAG-0447","socio74@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0447","",""],["PAG-0448","socio74@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0448","",""],["PAG-0449","socio74@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0449","",""],["PAG-0450","socio74@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0450","",""],["PAG-0451","socio75@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0451","",""],["PAG-0452","socio75@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0452","",""],["PAG-0453","socio75@futsalhaedo.com","2026-03",30000,"Pendiente","","",""],["PAG-0454","socio75@futsalhaedo.com","2026-04",30000,"Pendiente","","",""],["PAG-0455","socio75@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0456","socio75@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0457","socio76@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0457","",""],["PAG-0458","socio76@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0458","",""],["PAG-0459","socio76@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0459","",""],["PAG-0460","socio76@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0460","",""],["PAG-0461","socio76@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0461","",""],["PAG-0462","socio76@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0462","",""],["PAG-0463","socio77@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0463","",""],["PAG-0464","socio77@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0464","",""],["PAG-0465","socio77@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0465","",""],["PAG-0466","socio77@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0466","",""],["PAG-0467","socio77@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0467","",""],["PAG-0468","socio77@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0468","",""],["PAG-0469","socio78@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0469","",""],["PAG-0470","socio78@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0470","",""],["PAG-0471","socio78@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0471","",""],["PAG-0472","socio78@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0472","",""],["PAG-0473","socio78@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0473","",""],["PAG-0474","socio78@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0474","",""],["PAG-0475","socio79@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0475","",""],["PAG-0476","socio79@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0476","",""],["PAG-0477","socio79@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0477","",""],["PAG-0478","socio79@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0478","",""],["PAG-0479","socio79@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0479","",""],["PAG-0480","socio79@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0480","",""],["PAG-0481","socio80@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0481","",""],["PAG-0482","socio80@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0482","",""],["PAG-0483","socio80@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0483","",""],["PAG-0484","socio80@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0484","",""],["PAG-0485","socio80@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0486","socio80@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0487","socio81@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0487","",""],["PAG-0488","socio81@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0488","",""],["PAG-0489","socio81@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0489","",""],["PAG-0490","socio81@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0490","",""],["PAG-0491","socio81@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0491","",""],["PAG-0492","socio81@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0492","",""],["PAG-0493","socio82@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0493","",""],["PAG-0494","socio82@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0494","",""],["PAG-0495","socio82@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0495","",""],["PAG-0496","socio82@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0496","",""],["PAG-0497","socio82@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0497","",""],["PAG-0498","socio82@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0498","",""],["PAG-0499","socio83@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0499","",""],["PAG-0500","socio83@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0500","",""],["PAG-0501","socio83@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0501","",""],["PAG-0502","socio83@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0502","",""],["PAG-0503","socio83@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0503","",""],["PAG-0504","socio83@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0504","",""],["PAG-0505","socio84@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0505","",""],["PAG-0506","socio84@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0506","",""],["PAG-0507","socio84@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0507","",""],["PAG-0508","socio84@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0508","",""],["PAG-0509","socio84@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0509","",""],["PAG-0510","socio84@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0510","",""],["PAG-0511","socio85@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0511","",""],["PAG-0512","socio85@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0512","",""],["PAG-0513","socio85@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0513","",""],["PAG-0514","socio85@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0514","",""],["PAG-0515","socio85@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0515","",""],["PAG-0516","socio85@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0516","",""],["PAG-0517","socio86@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0517","",""],["PAG-0518","socio86@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0518","",""],["PAG-0519","socio86@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0519","",""],["PAG-0520","socio86@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0520","",""],["PAG-0521","socio86@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0521","",""],["PAG-0522","socio86@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0522","",""],["PAG-0523","socio87@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0523","",""],["PAG-0524","socio87@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0524","",""],["PAG-0525","socio87@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0525","",""],["PAG-0526","socio87@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0526","",""],["PAG-0527","socio87@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0527","",""],["PAG-0528","socio87@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0528","",""],["PAG-0529","socio88@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0529","",""],["PAG-0530","socio88@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0530","",""],["PAG-0531","socio88@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0531","",""],["PAG-0532","socio88@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0532","",""],["PAG-0533","socio88@futsalhaedo.com","2026-05",20000,"Pendiente","","",""],["PAG-0534","socio88@futsalhaedo.com","2026-06",20000,"Pendiente","","",""],["PAG-0535","socio89@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0535","",""],["PAG-0536","socio89@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0536","",""],["PAG-0537","socio89@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0537","",""],["PAG-0538","socio89@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0538","",""],["PAG-0539","socio89@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0539","",""],["PAG-0540","socio89@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0540","",""],["PAG-0541","socio90@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0541","",""],["PAG-0542","socio90@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0542","",""],["PAG-0543","socio90@futsalhaedo.com","2026-03",30000,"Pendiente","","",""],["PAG-0544","socio90@futsalhaedo.com","2026-04",30000,"Pendiente","","",""],["PAG-0545","socio90@futsalhaedo.com","2026-05",30000,"Pendiente","","",""],["PAG-0546","socio90@futsalhaedo.com","2026-06",30000,"Pendiente","","",""],["PAG-0547","socio91@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0547","",""],["PAG-0548","socio91@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0548","",""],["PAG-0549","socio91@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0549","",""],["PAG-0550","socio91@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0550","",""],["PAG-0551","socio91@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0551","",""],["PAG-0552","socio91@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0552","",""],["PAG-0553","socio92@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0553","",""],["PAG-0554","socio92@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0554","",""],["PAG-0555","socio92@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0555","",""],["PAG-0556","socio92@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0556","",""],["PAG-0557","socio92@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0557","",""],["PAG-0558","socio92@futsalhaedo.com","2026-06",15000,"Pendiente","","",""],["PAG-0559","socio93@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0559","",""],["PAG-0560","socio93@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0560","",""],["PAG-0561","socio93@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0561","",""],["PAG-0562","socio93@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0562","",""],["PAG-0563","socio93@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0563","",""],["PAG-0564","socio93@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0564","",""],["PAG-0565","socio94@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0565","",""],["PAG-0566","socio94@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0566","",""],["PAG-0567","socio94@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0567","",""],["PAG-0568","socio94@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0568","",""],["PAG-0569","socio94@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0569","",""],["PAG-0570","socio94@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0570","",""],["PAG-0571","socio95@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0571","",""],["PAG-0572","socio95@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0572","",""],["PAG-0573","socio95@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0573","",""],["PAG-0574","socio95@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0574","",""],["PAG-0575","socio95@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0575","",""],["PAG-0576","socio95@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0576","",""],["PAG-0577","socio96@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0577","",""],["PAG-0578","socio96@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0578","",""],["PAG-0579","socio96@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0579","",""],["PAG-0580","socio96@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0580","",""],["PAG-0581","socio96@futsalhaedo.com","2026-05",15000,"Pendiente","","",""],["PAG-0582","socio96@futsalhaedo.com","2026-06",15000,"Pendiente","","",""],["PAG-0583","socio97@futsalhaedo.com","2026-01",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0583","",""],["PAG-0584","socio97@futsalhaedo.com","2026-02",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0584","",""],["PAG-0585","socio97@futsalhaedo.com","2026-03",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0585","",""],["PAG-0586","socio97@futsalhaedo.com","2026-04",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0586","",""],["PAG-0587","socio97@futsalhaedo.com","2026-05",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0587","",""],["PAG-0588","socio97@futsalhaedo.com","2026-06",15000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0588","",""],["PAG-0589","socio98@futsalhaedo.com","2026-01",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0589","",""],["PAG-0590","socio98@futsalhaedo.com","2026-02",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0590","",""],["PAG-0591","socio98@futsalhaedo.com","2026-03",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0591","",""],["PAG-0592","socio98@futsalhaedo.com","2026-04",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0592","",""],["PAG-0593","socio98@futsalhaedo.com","2026-05",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0593","",""],["PAG-0594","socio98@futsalhaedo.com","2026-06",20000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0594","",""],["PAG-0595","socio99@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0595","",""],["PAG-0596","socio99@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0596","",""],["PAG-0597","socio99@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0597","",""],["PAG-0598","socio99@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0598","",""],["PAG-0599","socio99@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0599","",""],["PAG-0600","socio99@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0600","",""],["PAG-0601","socio100@futsalhaedo.com","2026-01",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0601","",""],["PAG-0602","socio100@futsalhaedo.com","2026-02",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0602","",""],["PAG-0603","socio100@futsalhaedo.com","2026-03",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0603","",""],["PAG-0604","socio100@futsalhaedo.com","2026-04",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0604","",""],["PAG-0605","socio100@futsalhaedo.com","2026-05",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0605","",""],["PAG-0606","socio100@futsalhaedo.com","2026-06",30000,"Pagado","https://www.mercadopago.com.ar/checkout/v1/redirect?pref_id=mock_pref_PAG-0606","",""],["PAG-0607","marielapmolina@gmail.com","2026-01",1000,"Pendiente","","",""],["PAG-0608","marielapmolina@gmail.com","2026-02",1000,"Pendiente","","",""],["PAG-0609","marielapmolina@gmail.com","2026-03",1000,"Pendiente","","",""],["PAG-0610","marielapmolina@gmail.com","2026-04",1000,"Pendiente","","",""],["PAG-0611","marielapmolina@gmail.com","2026-05",1000,"Pendiente","","",""],["PAG-0612","marielapmolina@gmail.com","2026-06",1000,"Pagado","","MercadoPago","2026-06-29"]];
  pagos.forEach(r => sheetPagos.appendRow(r));
  
  // 5. Configurar Hoja "Torneos"
  let sheetTorneos = ss.getSheetByName(HOJA_TORNEOS) || ss.insertSheet(HOJA_TORNEOS);
  sheetTorneos.clear();
  sheetTorneos.appendRow(["Torneo_ID","Name","Category"]);
  const torneos = [["T-001","Torneo EDEFI +35","Fut-May+35"],["T-002","Torneo EDEFI +42","Fut-May+42"],["T-003","Liga Femenina FutFEM","Fut-Fem-1ra"],["T-004","Liga Juvenil FUTSALA","Fut-Juv-5ta"],["T-005","Liga Baby EDEFI","Bab-2015"]];
  torneos.forEach(r => sheetTorneos.appendRow(r));
  
  // 6. Configurar Hoja "Finanzas_Torneos"
  let sheetFin = ss.getSheetByName(HOJA_FINANZAS) || ss.insertSheet(HOJA_FINANZAS);
  sheetFin.clear();
  sheetFin.appendRow(["Movimiento_ID","Torneo_ID","Type","Concept","Amount","Date","Payment_Method"]);
  const finanzas = [["MOV-0001","T-001","Gasto","Alquiler Cancha - Fecha 1 vs Huracán",8737,"2026-01-03","Efectivo"],["MOV-0002","T-001","Gasto","Arbitraje - Fecha 1 vs Huracán",5527,"2026-01-03","Efectivo"],["MOV-0003","T-001","Ingreso","Venta Entradas - Fecha 1 vs Huracán",12429,"2026-01-03","Efectivo"],["MOV-0004","T-001","Ingreso","Buffet/Cantina - Fecha 1 vs Huracán",8601,"2026-01-03","Efectivo"],["MOV-0005","T-001","Gasto","Traslado Delegación - Fecha 2 vs Morón",4455,"2026-01-10","Efectivo"],["MOV-0006","T-001","Ingreso","Rifas y Buffet - Fecha 2 vs Morón",5502,"2026-01-10","Efectivo"],["MOV-0007","T-001","Gasto","Alquiler Cancha - Fecha 3 vs Gimnasia",8181,"2026-01-17","Efectivo"],["MOV-0008","T-001","Gasto","Arbitraje - Fecha 3 vs Gimnasia",5570,"2026-01-17","Efectivo"],["MOV-0009","T-001","Ingreso","Venta Entradas - Fecha 3 vs Gimnasia",13595,"2026-01-17","Efectivo"],["MOV-0010","T-001","Ingreso","Buffet/Cantina - Fecha 3 vs Gimnasia",9004,"2026-01-17","Efectivo"],["MOV-0011","T-001","Gasto","Traslado Delegación - Fecha 4 vs Ramos Mejía",4257,"2026-01-24","Efectivo"],["MOV-0012","T-001","Ingreso","Rifas y Buffet - Fecha 4 vs Ramos Mejía",5619,"2026-01-24","Efectivo"],["MOV-0013","T-001","Gasto","Alquiler Cancha - Fecha 5 vs San Justo",8614,"2026-01-31","Efectivo"],["MOV-0014","T-001","Gasto","Arbitraje - Fecha 5 vs San Justo",5352,"2026-01-31","Efectivo"],["MOV-0015","T-001","Ingreso","Venta Entradas - Fecha 5 vs San Justo",14534,"2026-01-31","Efectivo"],["MOV-0016","T-001","Ingreso","Buffet/Cantina - Fecha 5 vs San Justo",9592,"2026-01-31","Efectivo"],["MOV-0017","T-001","Gasto","Traslado Delegación - Fecha 6 vs Ateneo",4173,"2026-02-07","Efectivo"],["MOV-0018","T-001","Ingreso","Rifas y Buffet - Fecha 6 vs Ateneo",6577,"2026-02-07","Efectivo"],["MOV-0019","T-001","Gasto","Alquiler Cancha - Fecha 7 vs Almafuerte",8413,"2026-02-14","Efectivo"],["MOV-0020","T-001","Gasto","Arbitraje - Fecha 7 vs Almafuerte",5599,"2026-02-14","Efectivo"],["MOV-0021","T-001","Ingreso","Venta Entradas - Fecha 7 vs Almafuerte",15707,"2026-02-14","Efectivo"],["MOV-0022","T-001","Ingreso","Buffet/Cantina - Fecha 7 vs Almafuerte",9771,"2026-02-14","Efectivo"],["MOV-0023","T-001","Gasto","Traslado Delegación - Fecha 8 vs Ituzaingó",4454,"2026-02-21","Efectivo"],["MOV-0024","T-001","Ingreso","Rifas y Buffet - Fecha 8 vs Ituzaingó",7622,"2026-02-21","Efectivo"],["MOV-0025","T-001","Gasto","Alquiler Cancha - Fecha 9 vs Liniers",8505,"2026-02-28","Efectivo"],["MOV-0026","T-001","Gasto","Arbitraje - Fecha 9 vs Liniers",5041,"2026-02-28","Efectivo"],["MOV-0027","T-001","Ingreso","Venta Entradas - Fecha 9 vs Liniers",16409,"2026-02-28","Efectivo"],["MOV-0028","T-001","Ingreso","Buffet/Cantina - Fecha 9 vs Liniers",10954,"2026-02-28","Efectivo"],["MOV-0029","T-001","Gasto","Traslado Delegación - Fecha 10 vs Almagro",4008,"2026-03-07","Efectivo"],["MOV-0030","T-001","Ingreso","Rifas y Buffet - Fecha 10 vs Almagro",7889,"2026-03-07","Efectivo"],["MOV-0031","T-001","Gasto","Alquiler Cancha - Fecha 11 vs Pearson",8858,"2026-03-14","Efectivo"],["MOV-0032","T-001","Gasto","Arbitraje - Fecha 11 vs Pearson",5335,"2026-03-14","Efectivo"],["MOV-0033","T-001","Ingreso","Venta Entradas - Fecha 11 vs Pearson",16423,"2026-03-14","Efectivo"],["MOV-0034","T-001","Ingreso","Buffet/Cantina - Fecha 11 vs Pearson",12934,"2026-03-14","Efectivo"],["MOV-0035","T-001","Gasto","Traslado Delegación - Fecha 12 vs CIDECO",4162,"2026-03-21","Efectivo"],["MOV-0036","T-001","Ingreso","Rifas y Buffet - Fecha 12 vs CIDECO",7841,"2026-03-21","Efectivo"],["MOV-0037","T-001","Gasto","Alquiler Cancha - Fecha 13 vs Estrella",8835,"2026-03-28","Efectivo"],["MOV-0038","T-001","Gasto","Arbitraje - Fecha 13 vs Estrella",5023,"2026-03-28","Efectivo"],["MOV-0039","T-001","Ingreso","Venta Entradas - Fecha 13 vs Estrella",18389,"2026-03-28","Efectivo"],["MOV-0040","T-001","Ingreso","Buffet/Cantina - Fecha 13 vs Estrella",12248,"2026-03-28","Efectivo"],["MOV-0041","T-001","Gasto","Traslado Delegación - Fecha 14 vs Defensores",4110,"2026-04-04","Efectivo"],["MOV-0042","T-001","Ingreso","Rifas y Buffet - Fecha 14 vs Defensores",9499,"2026-04-04","Efectivo"],["MOV-0043","T-001","Gasto","Alquiler Cancha - Fecha 15 vs Porteño",8140,"2026-04-11","Efectivo"],["MOV-0044","T-001","Gasto","Arbitraje - Fecha 15 vs Porteño",5007,"2026-04-11","Efectivo"],["MOV-0045","T-001","Ingreso","Venta Entradas - Fecha 15 vs Porteño",19354,"2026-04-11","Efectivo"],["MOV-0046","T-001","Ingreso","Buffet/Cantina - Fecha 15 vs Porteño",14360,"2026-04-11","Efectivo"],["MOV-0047","T-001","Gasto","Traslado Delegación - Fecha 16 vs Leloir",4311,"2026-04-18","Efectivo"],["MOV-0048","T-001","Ingreso","Rifas y Buffet - Fecha 16 vs Leloir",9601,"2026-04-18","Efectivo"],["MOV-0049","T-001","Gasto","Alquiler Cancha - Fecha 17 vs UAI",8223,"2026-04-25","Efectivo"],["MOV-0050","T-001","Gasto","Arbitraje - Fecha 17 vs UAI",5568,"2026-04-25","Efectivo"],["MOV-0051","T-001","Ingreso","Venta Entradas - Fecha 17 vs UAI",19297,"2026-04-25","Efectivo"],["MOV-0052","T-001","Ingreso","Buffet/Cantina - Fecha 17 vs UAI",13583,"2026-04-25","Efectivo"],["MOV-0053","T-001","Gasto","Traslado Delegación - Fecha 18 vs Pearson Juv",4191,"2026-05-02","Efectivo"],["MOV-0054","T-001","Ingreso","Rifas y Buffet - Fecha 18 vs Pearson Juv",10915,"2026-05-02","Efectivo"],["MOV-0055","T-001","Gasto","Alquiler Cancha - Fecha 19 vs Moron Futsal A",8072,"2026-05-09","Efectivo"],["MOV-0056","T-001","Gasto","Arbitraje - Fecha 19 vs Moron Futsal A",5217,"2026-05-09","Efectivo"],["MOV-0057","T-001","Ingreso","Venta Entradas - Fecha 19 vs Moron Futsal A",21172,"2026-05-09","Efectivo"],["MOV-0058","T-001","Ingreso","Buffet/Cantina - Fecha 19 vs Moron Futsal A",15098,"2026-05-09","Efectivo"],["MOV-0059","T-001","Gasto","Traslado Delegación - Fecha 20 vs Almafuerte Baby A",4120,"2026-05-16","Efectivo"],["MOV-0060","T-001","Ingreso","Rifas y Buffet - Fecha 20 vs Almafuerte Baby A",11215,"2026-05-16","Efectivo"],["MOV-0061","T-001","Gasto","Alquiler Cancha - Fecha 21 vs Lomas Futbol",8036,"2026-05-23","Efectivo"],["MOV-0062","T-001","Gasto","Arbitraje - Fecha 21 vs Lomas Futbol",5088,"2026-05-23","Efectivo"],["MOV-0063","T-001","Ingreso","Venta Entradas - Fecha 21 vs Lomas Futbol",21157,"2026-05-23","Efectivo"],["MOV-0064","T-001","Ingreso","Buffet/Cantina - Fecha 21 vs Lomas Futbol",15400,"2026-05-23","Efectivo"],["MOV-0065","T-001","Gasto","Traslado Delegación - Fecha 22 vs Porteño FutFem",4009,"2026-05-30","Efectivo"],["MOV-0066","T-001","Ingreso","Rifas y Buffet - Fecha 22 vs Porteño FutFem",10970,"2026-05-30","Efectivo"],["MOV-0067","T-001","Gasto","Alquiler Cancha - Fecha 23 vs Huracan San Justo Juv",8293,"2026-06-06","Efectivo"],["MOV-0068","T-001","Gasto","Arbitraje - Fecha 23 vs Huracan San Justo Juv",5058,"2026-06-06","Efectivo"],["MOV-0069","T-001","Ingreso","Venta Entradas - Fecha 23 vs Huracan San Justo Juv",22477,"2026-06-06","Efectivo"],["MOV-0070","T-001","Ingreso","Buffet/Cantina - Fecha 23 vs Huracan San Justo Juv",16356,"2026-06-06","Efectivo"],["MOV-0071","T-001","Gasto","Traslado Delegación - Fecha 24 vs Defensores Haedo Baby",4493,"2026-06-13","Efectivo"],["MOV-0072","T-001","Ingreso","Rifas y Buffet - Fecha 24 vs Defensores Haedo Baby",12304,"2026-06-13","Efectivo"],["MOV-0073","T-001","Gasto","Alquiler Cancha - Fecha 25 vs San Justo Baby",8043,"2026-06-20","Efectivo"],["MOV-0074","T-001","Gasto","Arbitraje - Fecha 25 vs San Justo Baby",5079,"2026-06-20","Efectivo"],["MOV-0075","T-001","Ingreso","Venta Entradas - Fecha 25 vs San Justo Baby",24959,"2026-06-20","Efectivo"],["MOV-0076","T-001","Ingreso","Buffet/Cantina - Fecha 25 vs San Justo Baby",17846,"2026-06-20","Efectivo"],["MOV-0077","T-001","Gasto","Traslado Delegación - Fecha 26 vs Huracán",4165,"2026-06-27","Efectivo"],["MOV-0078","T-001","Ingreso","Rifas y Buffet - Fecha 26 vs Huracán",12666,"2026-06-27","Efectivo"],["MOV-0079","T-002","Gasto","Traslado Delegación - Fecha 1 vs Morón",5069,"2026-01-03","Efectivo"],["MOV-0080","T-002","Ingreso","Rifas y Buffet - Fecha 1 vs Morón",6274,"2026-01-03","Efectivo"],["MOV-0081","T-002","Gasto","Alquiler Cancha - Fecha 2 vs Gimnasia",9755,"2026-01-10","Efectivo"],["MOV-0082","T-002","Gasto","Arbitraje - Fecha 2 vs Gimnasia",5835,"2026-01-10","Efectivo"],["MOV-0083","T-002","Ingreso","Venta Entradas - Fecha 2 vs Gimnasia",14027,"2026-01-10","Efectivo"],["MOV-0084","T-002","Ingreso","Buffet/Cantina - Fecha 2 vs Gimnasia",10768,"2026-01-10","Efectivo"],["MOV-0085","T-002","Gasto","Traslado Delegación - Fecha 3 vs Ramos Mejía",4998,"2026-01-17","Efectivo"],["MOV-0086","T-002","Ingreso","Rifas y Buffet - Fecha 3 vs Ramos Mejía",6303,"2026-01-17","Efectivo"],["MOV-0087","T-002","Gasto","Alquiler Cancha - Fecha 4 vs San Justo",9412,"2026-01-24","Efectivo"],["MOV-0088","T-002","Gasto","Arbitraje - Fecha 4 vs San Justo",5970,"2026-01-24","Efectivo"],["MOV-0089","T-002","Ingreso","Venta Entradas - Fecha 4 vs San Justo",14932,"2026-01-24","Efectivo"],["MOV-0090","T-002","Ingreso","Buffet/Cantina - Fecha 4 vs San Justo",11062,"2026-01-24","Efectivo"],["MOV-0091","T-002","Gasto","Traslado Delegación - Fecha 5 vs Ateneo",5033,"2026-01-31","Efectivo"],["MOV-0092","T-002","Ingreso","Rifas y Buffet - Fecha 5 vs Ateneo",7273,"2026-01-31","Efectivo"],["MOV-0093","T-002","Gasto","Alquiler Cancha - Fecha 6 vs Almafuerte",9554,"2026-02-07","Efectivo"],["MOV-0094","T-002","Gasto","Arbitraje - Fecha 6 vs Almafuerte",5451,"2026-02-07","Efectivo"],["MOV-0095","T-002","Ingreso","Venta Entradas - Fecha 6 vs Almafuerte",15459,"2026-02-07","Efectivo"],["MOV-0096","T-002","Ingreso","Buffet/Cantina - Fecha 6 vs Almafuerte",11573,"2026-02-07","Efectivo"],["MOV-0097","T-002","Gasto","Traslado Delegación - Fecha 7 vs Ituzaingó",4965,"2026-02-14","Efectivo"],["MOV-0098","T-002","Ingreso","Rifas y Buffet - Fecha 7 vs Ituzaingó",7418,"2026-02-14","Efectivo"],["MOV-0099","T-002","Gasto","Alquiler Cancha - Fecha 8 vs Liniers",9679,"2026-02-21","Efectivo"],["MOV-0100","T-002","Gasto","Arbitraje - Fecha 8 vs Liniers",5547,"2026-02-21","Efectivo"],["MOV-0101","T-002","Ingreso","Venta Entradas - Fecha 8 vs Liniers",16764,"2026-02-21","Efectivo"],["MOV-0102","T-002","Ingreso","Buffet/Cantina - Fecha 8 vs Liniers",13034,"2026-02-21","Efectivo"],["MOV-0103","T-002","Gasto","Traslado Delegación - Fecha 9 vs Almagro",5014,"2026-02-28","Efectivo"],["MOV-0104","T-002","Ingreso","Rifas y Buffet - Fecha 9 vs Almagro",8445,"2026-02-28","Efectivo"],["MOV-0105","T-002","Gasto","Alquiler Cancha - Fecha 10 vs Pearson",9798,"2026-03-07","Efectivo"],["MOV-0106","T-002","Gasto","Arbitraje - Fecha 10 vs Pearson",5883,"2026-03-07","Efectivo"],["MOV-0107","T-002","Ingreso","Venta Entradas - Fecha 10 vs Pearson",17762,"2026-03-07","Efectivo"],["MOV-0108","T-002","Ingreso","Buffet/Cantina - Fecha 10 vs Pearson",12113,"2026-03-07","Efectivo"],["MOV-0109","T-002","Gasto","Traslado Delegación - Fecha 11 vs CIDECO",4986,"2026-03-14","Efectivo"],["MOV-0110","T-002","Ingreso","Rifas y Buffet - Fecha 11 vs CIDECO",9017,"2026-03-14","Efectivo"],["MOV-0111","T-002","Gasto","Alquiler Cancha - Fecha 12 vs Estrella",8905,"2026-03-21","Efectivo"],["MOV-0112","T-002","Gasto","Arbitraje - Fecha 12 vs Estrella",5707,"2026-03-21","Efectivo"],["MOV-0113","T-002","Ingreso","Venta Entradas - Fecha 12 vs Estrella",18761,"2026-03-21","Efectivo"],["MOV-0114","T-002","Ingreso","Buffet/Cantina - Fecha 12 vs Estrella",13497,"2026-03-21","Efectivo"],["MOV-0115","T-002","Gasto","Traslado Delegación - Fecha 13 vs Defensores",4984,"2026-03-28","Efectivo"],["MOV-0116","T-002","Ingreso","Rifas y Buffet - Fecha 13 vs Defensores",10503,"2026-03-28","Efectivo"],["MOV-0117","T-002","Gasto","Alquiler Cancha - Fecha 14 vs Porteño",8863,"2026-04-04","Efectivo"],["MOV-0118","T-002","Gasto","Arbitraje - Fecha 14 vs Porteño",5418,"2026-04-04","Efectivo"],["MOV-0119","T-002","Ingreso","Venta Entradas - Fecha 14 vs Porteño",18932,"2026-04-04","Efectivo"],["MOV-0120","T-002","Ingreso","Buffet/Cantina - Fecha 14 vs Porteño",15289,"2026-04-04","Efectivo"],["MOV-0121","T-002","Gasto","Traslado Delegación - Fecha 15 vs Leloir",4880,"2026-04-11","Efectivo"],["MOV-0122","T-002","Ingreso","Rifas y Buffet - Fecha 15 vs Leloir",10788,"2026-04-11","Efectivo"],["MOV-0123","T-002","Gasto","Alquiler Cancha - Fecha 16 vs UAI",9637,"2026-04-18","Efectivo"],["MOV-0124","T-002","Gasto","Arbitraje - Fecha 16 vs UAI",5548,"2026-04-18","Efectivo"],["MOV-0125","T-002","Ingreso","Venta Entradas - Fecha 16 vs UAI",20343,"2026-04-18","Efectivo"],["MOV-0126","T-002","Ingreso","Buffet/Cantina - Fecha 16 vs UAI",15439,"2026-04-18","Efectivo"],["MOV-0127","T-002","Gasto","Traslado Delegación - Fecha 17 vs Pearson Juv",4972,"2026-04-25","Efectivo"],["MOV-0128","T-002","Ingreso","Rifas y Buffet - Fecha 17 vs Pearson Juv",10904,"2026-04-25","Efectivo"],["MOV-0129","T-002","Gasto","Alquiler Cancha - Fecha 18 vs Moron Futsal A",9328,"2026-05-02","Efectivo"],["MOV-0130","T-002","Gasto","Arbitraje - Fecha 18 vs Moron Futsal A",5996,"2026-05-02","Efectivo"],["MOV-0131","T-002","Ingreso","Venta Entradas - Fecha 18 vs Moron Futsal A",21771,"2026-05-02","Efectivo"],["MOV-0132","T-002","Ingreso","Buffet/Cantina - Fecha 18 vs Moron Futsal A",16214,"2026-05-02","Efectivo"],["MOV-0133","T-002","Gasto","Traslado Delegación - Fecha 19 vs Almafuerte Baby A",5098,"2026-05-09","Efectivo"],["MOV-0134","T-002","Ingreso","Rifas y Buffet - Fecha 19 vs Almafuerte Baby A",12060,"2026-05-09","Efectivo"],["MOV-0135","T-002","Gasto","Alquiler Cancha - Fecha 20 vs Lomas Futbol",9797,"2026-05-16","Efectivo"],["MOV-0136","T-002","Gasto","Arbitraje - Fecha 20 vs Lomas Futbol",5909,"2026-05-16","Efectivo"],["MOV-0137","T-002","Ingreso","Venta Entradas - Fecha 20 vs Lomas Futbol",23273,"2026-05-16","Efectivo"],["MOV-0138","T-002","Ingreso","Buffet/Cantina - Fecha 20 vs Lomas Futbol",17984,"2026-05-16","Efectivo"],["MOV-0139","T-002","Gasto","Traslado Delegación - Fecha 21 vs Porteño FutFem",4711,"2026-05-23","Efectivo"],["MOV-0140","T-002","Ingreso","Rifas y Buffet - Fecha 21 vs Porteño FutFem",11637,"2026-05-23","Efectivo"],["MOV-0141","T-002","Gasto","Alquiler Cancha - Fecha 22 vs Huracan San Justo Juv",9528,"2026-05-30","Efectivo"],["MOV-0142","T-002","Gasto","Arbitraje - Fecha 22 vs Huracan San Justo Juv",5427,"2026-05-30","Efectivo"],["MOV-0143","T-002","Ingreso","Venta Entradas - Fecha 22 vs Huracan San Justo Juv",24475,"2026-05-30","Efectivo"],["MOV-0144","T-002","Ingreso","Buffet/Cantina - Fecha 22 vs Huracan San Justo Juv",18412,"2026-05-30","Efectivo"],["MOV-0145","T-002","Gasto","Traslado Delegación - Fecha 23 vs Defensores Haedo Baby",4720,"2026-06-06","Efectivo"],["MOV-0146","T-002","Ingreso","Rifas y Buffet - Fecha 23 vs Defensores Haedo Baby",12696,"2026-06-06","Efectivo"],["MOV-0147","T-002","Gasto","Alquiler Cancha - Fecha 24 vs San Justo Baby",9353,"2026-06-13","Efectivo"],["MOV-0148","T-002","Gasto","Arbitraje - Fecha 24 vs San Justo Baby",5970,"2026-06-13","Efectivo"],["MOV-0149","T-002","Ingreso","Venta Entradas - Fecha 24 vs San Justo Baby",23687,"2026-06-13","Efectivo"],["MOV-0150","T-002","Ingreso","Buffet/Cantina - Fecha 24 vs San Justo Baby",19016,"2026-06-13","Efectivo"],["MOV-0151","T-002","Gasto","Traslado Delegación - Fecha 25 vs Huracán",5070,"2026-06-20","Efectivo"],["MOV-0152","T-002","Ingreso","Rifas y Buffet - Fecha 25 vs Huracán",14193,"2026-06-20","Efectivo"],["MOV-0153","T-002","Gasto","Alquiler Cancha - Fecha 26 vs Morón",9788,"2026-06-27","Efectivo"],["MOV-0154","T-002","Gasto","Arbitraje - Fecha 26 vs Morón",5442,"2026-06-27","Efectivo"],["MOV-0155","T-002","Ingreso","Venta Entradas - Fecha 26 vs Morón",26308,"2026-06-27","Efectivo"],["MOV-0156","T-002","Ingreso","Buffet/Cantina - Fecha 26 vs Morón",20258,"2026-06-27","Efectivo"],["MOV-0157","T-003","Gasto","Alquiler Cancha - Fecha 1 vs Gimnasia Fem",10475,"2026-01-03","Efectivo"],["MOV-0158","T-003","Gasto","Arbitraje - Fecha 1 vs Gimnasia Fem",6165,"2026-01-03","Efectivo"],["MOV-0159","T-003","Ingreso","Venta Entradas - Fecha 1 vs Gimnasia Fem",14898,"2026-01-03","Efectivo"],["MOV-0160","T-003","Ingreso","Buffet/Cantina - Fecha 1 vs Gimnasia Fem",11553,"2026-01-03","Efectivo"],["MOV-0161","T-003","Gasto","Traslado Delegación - Fecha 2 vs Ramos Mejía Fem",5439,"2026-01-10","Efectivo"],["MOV-0162","T-003","Ingreso","Rifas y Buffet - Fecha 2 vs Ramos Mejía Fem",8145,"2026-01-10","Efectivo"],["MOV-0163","T-003","Gasto","Alquiler Cancha - Fecha 3 vs San Justo Fem",9900,"2026-01-17","Efectivo"],["MOV-0164","T-003","Gasto","Arbitraje - Fecha 3 vs San Justo Fem",6197,"2026-01-17","Efectivo"],["MOV-0165","T-003","Ingreso","Venta Entradas - Fecha 3 vs San Justo Fem",14936,"2026-01-17","Efectivo"],["MOV-0166","T-003","Ingreso","Buffet/Cantina - Fecha 3 vs San Justo Fem",11952,"2026-01-17","Efectivo"],["MOV-0167","T-003","Gasto","Traslado Delegación - Fecha 4 vs Ateneo Fem",5402,"2026-01-24","Efectivo"],["MOV-0168","T-003","Ingreso","Rifas y Buffet - Fecha 4 vs Ateneo Fem",7654,"2026-01-24","Efectivo"],["MOV-0169","T-003","Gasto","Alquiler Cancha - Fecha 5 vs Almafuerte Fem",9846,"2026-01-31","Efectivo"],["MOV-0170","T-003","Gasto","Arbitraje - Fecha 5 vs Almafuerte Fem",6333,"2026-01-31","Efectivo"],["MOV-0171","T-003","Ingreso","Venta Entradas - Fecha 5 vs Almafuerte Fem",15710,"2026-01-31","Efectivo"],["MOV-0172","T-003","Ingreso","Buffet/Cantina - Fecha 5 vs Almafuerte Fem",12279,"2026-01-31","Efectivo"],["MOV-0173","T-003","Gasto","Traslado Delegación - Fecha 6 vs Ituzaingó Fem",5639,"2026-02-07","Efectivo"],["MOV-0174","T-003","Ingreso","Rifas y Buffet - Fecha 6 vs Ituzaingó Fem",8199,"2026-02-07","Efectivo"],["MOV-0175","T-003","Gasto","Alquiler Cancha - Fecha 7 vs Liniers Fem",9857,"2026-02-14","Efectivo"],["MOV-0176","T-003","Gasto","Arbitraje - Fecha 7 vs Liniers Fem",6306,"2026-02-14","Efectivo"],["MOV-0177","T-003","Ingreso","Venta Entradas - Fecha 7 vs Liniers Fem",17098,"2026-02-14","Efectivo"],["MOV-0178","T-003","Ingreso","Buffet/Cantina - Fecha 7 vs Liniers Fem",12589,"2026-02-14","Efectivo"],["MOV-0179","T-003","Gasto","Traslado Delegación - Fecha 8 vs Almagro Fem",5455,"2026-02-21","Efectivo"],["MOV-0180","T-003","Ingreso","Rifas y Buffet - Fecha 8 vs Almagro Fem",8765,"2026-02-21","Efectivo"],["MOV-0181","T-003","Gasto","Alquiler Cancha - Fecha 9 vs Pearson Fem",9759,"2026-02-28","Efectivo"],["MOV-0182","T-003","Gasto","Arbitraje - Fecha 9 vs Pearson Fem",6309,"2026-02-28","Efectivo"],["MOV-0183","T-003","Ingreso","Venta Entradas - Fecha 9 vs Pearson Fem",18810,"2026-02-28","Efectivo"],["MOV-0184","T-003","Ingreso","Buffet/Cantina - Fecha 9 vs Pearson Fem",15194,"2026-02-28","Efectivo"],["MOV-0185","T-003","Gasto","Traslado Delegación - Fecha 10 vs CIDECO Fem",5315,"2026-03-07","Efectivo"],["MOV-0186","T-003","Ingreso","Rifas y Buffet - Fecha 10 vs CIDECO Fem",10259,"2026-03-07","Efectivo"],["MOV-0187","T-003","Gasto","Alquiler Cancha - Fecha 11 vs Estrella Fem",9847,"2026-03-14","Efectivo"],["MOV-0188","T-003","Gasto","Arbitraje - Fecha 11 vs Estrella Fem",6132,"2026-03-14","Efectivo"],["MOV-0189","T-003","Ingreso","Venta Entradas - Fecha 11 vs Estrella Fem",18672,"2026-03-14","Efectivo"],["MOV-0190","T-003","Ingreso","Buffet/Cantina - Fecha 11 vs Estrella Fem",14274,"2026-03-14","Efectivo"],["MOV-0191","T-003","Gasto","Traslado Delegación - Fecha 12 vs Defensores Fem",5555,"2026-03-21","Efectivo"],["MOV-0192","T-003","Ingreso","Rifas y Buffet - Fecha 12 vs Defensores Fem",9845,"2026-03-21","Efectivo"],["MOV-0193","T-003","Gasto","Alquiler Cancha - Fecha 13 vs Porteño Fem",9705,"2026-03-28","Efectivo"],["MOV-0194","T-003","Gasto","Arbitraje - Fecha 13 vs Porteño Fem",6047,"2026-03-28","Efectivo"],["MOV-0195","T-003","Ingreso","Venta Entradas - Fecha 13 vs Porteño Fem",19727,"2026-03-28","Efectivo"],["MOV-0196","T-003","Ingreso","Buffet/Cantina - Fecha 13 vs Porteño Fem",15953,"2026-03-28","Efectivo"],["MOV-0197","T-003","Gasto","Traslado Delegación - Fecha 14 vs Leloir Fem",5422,"2026-04-04","Efectivo"],["MOV-0198","T-003","Ingreso","Rifas y Buffet - Fecha 14 vs Leloir Fem",10634,"2026-04-04","Efectivo"],["MOV-0199","T-003","Gasto","Alquiler Cancha - Fecha 15 vs UAI Fem",9752,"2026-04-11","Efectivo"],["MOV-0200","T-003","Gasto","Arbitraje - Fecha 15 vs UAI Fem",6314,"2026-04-11","Efectivo"],["MOV-0201","T-003","Ingreso","Venta Entradas - Fecha 15 vs UAI Fem",20701,"2026-04-11","Efectivo"],["MOV-0202","T-003","Ingreso","Buffet/Cantina - Fecha 15 vs UAI Fem",16932,"2026-04-11","Efectivo"],["MOV-0203","T-003","Gasto","Traslado Delegación - Fecha 16 vs Pearson Juv Fem",5418,"2026-04-18","Efectivo"],["MOV-0204","T-003","Ingreso","Rifas y Buffet - Fecha 16 vs Pearson Juv Fem",11306,"2026-04-18","Efectivo"],["MOV-0205","T-003","Gasto","Alquiler Cancha - Fecha 17 vs Moron Futsal A Fem",9928,"2026-04-25","Efectivo"],["MOV-0206","T-003","Gasto","Arbitraje - Fecha 17 vs Moron Futsal A Fem",5913,"2026-04-25","Efectivo"],["MOV-0207","T-003","Ingreso","Venta Entradas - Fecha 17 vs Moron Futsal A Fem",21896,"2026-04-25","Efectivo"],["MOV-0208","T-003","Ingreso","Buffet/Cantina - Fecha 17 vs Moron Futsal A Fem",16564,"2026-04-25","Efectivo"],["MOV-0209","T-003","Gasto","Traslado Delegación - Fecha 18 vs Almafuerte Baby A Fem",5370,"2026-05-02","Efectivo"],["MOV-0210","T-003","Ingreso","Rifas y Buffet - Fecha 18 vs Almafuerte Baby A Fem",12087,"2026-05-02","Efectivo"],["MOV-0211","T-003","Gasto","Alquiler Cancha - Fecha 19 vs Lomas Futbol Fem",10102,"2026-05-09","Efectivo"],["MOV-0212","T-003","Gasto","Arbitraje - Fecha 19 vs Lomas Futbol Fem",6391,"2026-05-09","Efectivo"],["MOV-0213","T-003","Ingreso","Venta Entradas - Fecha 19 vs Lomas Futbol Fem",23990,"2026-05-09","Efectivo"],["MOV-0214","T-003","Ingreso","Buffet/Cantina - Fecha 19 vs Lomas Futbol Fem",19130,"2026-05-09","Efectivo"],["MOV-0215","T-003","Gasto","Traslado Delegación - Fecha 20 vs Porteño FutFem Fem",5602,"2026-05-16","Efectivo"],["MOV-0216","T-003","Ingreso","Rifas y Buffet - Fecha 20 vs Porteño FutFem Fem",13086,"2026-05-16","Efectivo"],["MOV-0217","T-003","Gasto","Alquiler Cancha - Fecha 21 vs Huracan San Justo Juv Fem",10242,"2026-05-23","Efectivo"],["MOV-0218","T-003","Gasto","Arbitraje - Fecha 21 vs Huracan San Justo Juv Fem",5841,"2026-05-23","Efectivo"],["MOV-0219","T-003","Ingreso","Venta Entradas - Fecha 21 vs Huracan San Justo Juv Fem",24205,"2026-05-23","Efectivo"],["MOV-0220","T-003","Ingreso","Buffet/Cantina - Fecha 21 vs Huracan San Justo Juv Fem",18344,"2026-05-23","Efectivo"],["MOV-0221","T-003","Gasto","Traslado Delegación - Fecha 22 vs Defensores Haedo Baby Fem",5679,"2026-05-30","Efectivo"],["MOV-0222","T-003","Ingreso","Rifas y Buffet - Fecha 22 vs Defensores Haedo Baby Fem",13398,"2026-05-30","Efectivo"],["MOV-0223","T-003","Gasto","Alquiler Cancha - Fecha 23 vs San Justo Baby Fem",10516,"2026-06-06","Efectivo"],["MOV-0224","T-003","Gasto","Arbitraje - Fecha 23 vs San Justo Baby Fem",5988,"2026-06-06","Efectivo"],["MOV-0225","T-003","Ingreso","Venta Entradas - Fecha 23 vs San Justo Baby Fem",25353,"2026-06-06","Efectivo"],["MOV-0226","T-003","Ingreso","Buffet/Cantina - Fecha 23 vs San Justo Baby Fem",19836,"2026-06-06","Efectivo"],["MOV-0227","T-003","Gasto","Traslado Delegación - Fecha 24 vs Huracán Fem",5475,"2026-06-13","Efectivo"],["MOV-0228","T-003","Ingreso","Rifas y Buffet - Fecha 24 vs Huracán Fem",14421,"2026-06-13","Efectivo"],["MOV-0229","T-003","Gasto","Alquiler Cancha - Fecha 25 vs Morón Fem",9715,"2026-06-20","Efectivo"],["MOV-0230","T-003","Gasto","Arbitraje - Fecha 25 vs Morón Fem",6260,"2026-06-20","Efectivo"],["MOV-0231","T-003","Ingreso","Venta Entradas - Fecha 25 vs Morón Fem",26205,"2026-06-20","Efectivo"],["MOV-0232","T-003","Ingreso","Buffet/Cantina - Fecha 25 vs Morón Fem",19948,"2026-06-20","Efectivo"],["MOV-0233","T-003","Gasto","Traslado Delegación - Fecha 26 vs Gimnasia Fem",5294,"2026-06-27","Efectivo"],["MOV-0234","T-003","Ingreso","Rifas y Buffet - Fecha 26 vs Gimnasia Fem",14370,"2026-06-27","Efectivo"],["MOV-0235","T-004","Gasto","Traslado Delegación - Fecha 1 vs Ramos Mejía Juv",6259,"2026-01-03","Efectivo"],["MOV-0236","T-004","Ingreso","Rifas y Buffet - Fecha 1 vs Ramos Mejía Juv",8928,"2026-01-03","Efectivo"],["MOV-0237","T-004","Gasto","Alquiler Cancha - Fecha 2 vs San Justo Juv",11122,"2026-01-10","Efectivo"],["MOV-0238","T-004","Gasto","Arbitraje - Fecha 2 vs San Justo Juv",6544,"2026-01-10","Efectivo"],["MOV-0239","T-004","Ingreso","Venta Entradas - Fecha 2 vs San Justo Juv",15591,"2026-01-10","Efectivo"],["MOV-0240","T-004","Ingreso","Buffet/Cantina - Fecha 2 vs San Justo Juv",13135,"2026-01-10","Efectivo"],["MOV-0241","T-004","Gasto","Traslado Delegación - Fecha 3 vs Ateneo Juv",6255,"2026-01-17","Efectivo"],["MOV-0242","T-004","Ingreso","Rifas y Buffet - Fecha 3 vs Ateneo Juv",9161,"2026-01-17","Efectivo"],["MOV-0243","T-004","Gasto","Alquiler Cancha - Fecha 4 vs Almafuerte Juv",10616,"2026-01-24","Efectivo"],["MOV-0244","T-004","Gasto","Arbitraje - Fecha 4 vs Almafuerte Juv",6764,"2026-01-24","Efectivo"],["MOV-0245","T-004","Ingreso","Venta Entradas - Fecha 4 vs Almafuerte Juv",17306,"2026-01-24","Efectivo"],["MOV-0246","T-004","Ingreso","Buffet/Cantina - Fecha 4 vs Almafuerte Juv",14039,"2026-01-24","Efectivo"],["MOV-0247","T-004","Gasto","Traslado Delegación - Fecha 5 vs Ituzaingó Juv",6063,"2026-01-31","Efectivo"],["MOV-0248","T-004","Ingreso","Rifas y Buffet - Fecha 5 vs Ituzaingó Juv",9784,"2026-01-31","Efectivo"],["MOV-0249","T-004","Gasto","Alquiler Cancha - Fecha 6 vs Liniers Juv",10604,"2026-02-07","Efectivo"],["MOV-0250","T-004","Gasto","Arbitraje - Fecha 6 vs Liniers Juv",6490,"2026-02-07","Efectivo"],["MOV-0251","T-004","Ingreso","Venta Entradas - Fecha 6 vs Liniers Juv",16693,"2026-02-07","Efectivo"],["MOV-0252","T-004","Ingreso","Buffet/Cantina - Fecha 6 vs Liniers Juv",15170,"2026-02-07","Efectivo"],["MOV-0253","T-004","Gasto","Traslado Delegación - Fecha 7 vs Almagro Juv",5963,"2026-02-14","Efectivo"],["MOV-0254","T-004","Ingreso","Rifas y Buffet - Fecha 7 vs Almagro Juv",9917,"2026-02-14","Efectivo"],["MOV-0255","T-004","Gasto","Alquiler Cancha - Fecha 8 vs Pearson Juv",10614,"2026-02-21","Efectivo"],["MOV-0256","T-004","Gasto","Arbitraje - Fecha 8 vs Pearson Juv",6430,"2026-02-21","Efectivo"],["MOV-0257","T-004","Ingreso","Venta Entradas - Fecha 8 vs Pearson Juv",18176,"2026-02-21","Efectivo"],["MOV-0258","T-004","Ingreso","Buffet/Cantina - Fecha 8 vs Pearson Juv",16296,"2026-02-21","Efectivo"],["MOV-0259","T-004","Gasto","Traslado Delegación - Fecha 9 vs CIDECO Juv",5910,"2026-02-28","Efectivo"],["MOV-0260","T-004","Ingreso","Rifas y Buffet - Fecha 9 vs CIDECO Juv",11271,"2026-02-28","Efectivo"],["MOV-0261","T-004","Gasto","Alquiler Cancha - Fecha 10 vs Estrella Juv",10589,"2026-03-07","Efectivo"],["MOV-0262","T-004","Gasto","Arbitraje - Fecha 10 vs Estrella Juv",6265,"2026-03-07","Efectivo"],["MOV-0263","T-004","Ingreso","Venta Entradas - Fecha 10 vs Estrella Juv",20343,"2026-03-07","Efectivo"],["MOV-0264","T-004","Ingreso","Buffet/Cantina - Fecha 10 vs Estrella Juv",16278,"2026-03-07","Efectivo"],["MOV-0265","T-004","Gasto","Traslado Delegación - Fecha 11 vs Defensores Juv",6136,"2026-03-14","Efectivo"],["MOV-0266","T-004","Ingreso","Rifas y Buffet - Fecha 11 vs Defensores Juv",10929,"2026-03-14","Efectivo"],["MOV-0267","T-004","Gasto","Alquiler Cancha - Fecha 12 vs Porteño Juv",10494,"2026-03-21","Efectivo"],["MOV-0268","T-004","Gasto","Arbitraje - Fecha 12 vs Porteño Juv",6449,"2026-03-21","Efectivo"],["MOV-0269","T-004","Ingreso","Venta Entradas - Fecha 12 vs Porteño Juv",21319,"2026-03-21","Efectivo"],["MOV-0270","T-004","Ingreso","Buffet/Cantina - Fecha 12 vs Porteño Juv",16663,"2026-03-21","Efectivo"],["MOV-0271","T-004","Gasto","Traslado Delegación - Fecha 13 vs Leloir Juv",5844,"2026-03-28","Efectivo"],["MOV-0272","T-004","Ingreso","Rifas y Buffet - Fecha 13 vs Leloir Juv",12415,"2026-03-28","Efectivo"],["MOV-0273","T-004","Gasto","Alquiler Cancha - Fecha 14 vs UAI Juv",11220,"2026-04-04","Efectivo"],["MOV-0274","T-004","Gasto","Arbitraje - Fecha 14 vs UAI Juv",6429,"2026-04-04","Efectivo"],["MOV-0275","T-004","Ingreso","Venta Entradas - Fecha 14 vs UAI Juv",22133,"2026-04-04","Efectivo"],["MOV-0276","T-004","Ingreso","Buffet/Cantina - Fecha 14 vs UAI Juv",16720,"2026-04-04","Efectivo"],["MOV-0277","T-004","Gasto","Traslado Delegación - Fecha 15 vs Pearson Juv Juv",6040,"2026-04-11","Efectivo"],["MOV-0278","T-004","Ingreso","Rifas y Buffet - Fecha 15 vs Pearson Juv Juv",12635,"2026-04-11","Efectivo"],["MOV-0279","T-004","Gasto","Alquiler Cancha - Fecha 16 vs Moron Futsal A Juv",11171,"2026-04-18","Efectivo"],["MOV-0280","T-004","Gasto","Arbitraje - Fecha 16 vs Moron Futsal A Juv",6476,"2026-04-18","Efectivo"],["MOV-0281","T-004","Ingreso","Venta Entradas - Fecha 16 vs Moron Futsal A Juv",23432,"2026-04-18","Efectivo"],["MOV-0282","T-004","Ingreso","Buffet/Cantina - Fecha 16 vs Moron Futsal A Juv",19353,"2026-04-18","Efectivo"],["MOV-0283","T-004","Gasto","Traslado Delegación - Fecha 17 vs Almafuerte Baby A Juv",6208,"2026-04-25","Efectivo"],["MOV-0284","T-004","Ingreso","Rifas y Buffet - Fecha 17 vs Almafuerte Baby A Juv",12674,"2026-04-25","Efectivo"],["MOV-0285","T-004","Gasto","Alquiler Cancha - Fecha 18 vs Lomas Futbol Juv",10412,"2026-05-02","Efectivo"],["MOV-0286","T-004","Gasto","Arbitraje - Fecha 18 vs Lomas Futbol Juv",6781,"2026-05-02","Efectivo"],["MOV-0287","T-004","Ingreso","Venta Entradas - Fecha 18 vs Lomas Futbol Juv",23009,"2026-05-02","Efectivo"],["MOV-0288","T-004","Ingreso","Buffet/Cantina - Fecha 18 vs Lomas Futbol Juv",18488,"2026-05-02","Efectivo"],["MOV-0289","T-004","Gasto","Traslado Delegación - Fecha 19 vs Porteño FutFem Juv",6220,"2026-05-09","Efectivo"],["MOV-0290","T-004","Ingreso","Rifas y Buffet - Fecha 19 vs Porteño FutFem Juv",13130,"2026-05-09","Efectivo"],["MOV-0291","T-004","Gasto","Alquiler Cancha - Fecha 20 vs Huracan San Justo Juv Juv",11149,"2026-05-16","Efectivo"],["MOV-0292","T-004","Gasto","Arbitraje - Fecha 20 vs Huracan San Justo Juv Juv",6718,"2026-05-16","Efectivo"],["MOV-0293","T-004","Ingreso","Venta Entradas - Fecha 20 vs Huracan San Justo Juv Juv",23643,"2026-05-16","Efectivo"],["MOV-0294","T-004","Ingreso","Buffet/Cantina - Fecha 20 vs Huracan San Justo Juv Juv",19596,"2026-05-16","Efectivo"],["MOV-0295","T-004","Gasto","Traslado Delegación - Fecha 21 vs Defensores Haedo Baby Juv",6180,"2026-05-23","Efectivo"],["MOV-0296","T-004","Ingreso","Rifas y Buffet - Fecha 21 vs Defensores Haedo Baby Juv",14169,"2026-05-23","Efectivo"],["MOV-0297","T-004","Gasto","Alquiler Cancha - Fecha 22 vs San Justo Baby Juv",11000,"2026-05-30","Efectivo"],["MOV-0298","T-004","Gasto","Arbitraje - Fecha 22 vs San Justo Baby Juv",6624,"2026-05-30","Efectivo"],["MOV-0299","T-004","Ingreso","Venta Entradas - Fecha 22 vs San Justo Baby Juv",25629,"2026-05-30","Efectivo"],["MOV-0300","T-004","Ingreso","Buffet/Cantina - Fecha 22 vs San Justo Baby Juv",20913,"2026-05-30","Efectivo"],["MOV-0301","T-004","Gasto","Traslado Delegación - Fecha 23 vs Huracán Juv",6235,"2026-06-06","Efectivo"],["MOV-0302","T-004","Ingreso","Rifas y Buffet - Fecha 23 vs Huracán Juv",14522,"2026-06-06","Efectivo"],["MOV-0303","T-004","Gasto","Alquiler Cancha - Fecha 24 vs Morón Juv",11337,"2026-06-13","Efectivo"],["MOV-0304","T-004","Gasto","Arbitraje - Fecha 24 vs Morón Juv",6796,"2026-06-13","Efectivo"],["MOV-0305","T-004","Ingreso","Venta Entradas - Fecha 24 vs Morón Juv",25849,"2026-06-13","Efectivo"],["MOV-0306","T-004","Ingreso","Buffet/Cantina - Fecha 24 vs Morón Juv",21696,"2026-06-13","Efectivo"],["MOV-0307","T-004","Gasto","Traslado Delegación - Fecha 25 vs Gimnasia Juv",5801,"2026-06-20","Efectivo"],["MOV-0308","T-004","Ingreso","Rifas y Buffet - Fecha 25 vs Gimnasia Juv",15197,"2026-06-20","Efectivo"],["MOV-0309","T-004","Gasto","Alquiler Cancha - Fecha 26 vs Ramos Mejía Juv",10797,"2026-06-27","Efectivo"],["MOV-0310","T-004","Gasto","Arbitraje - Fecha 26 vs Ramos Mejía Juv",6755,"2026-06-27","Efectivo"],["MOV-0311","T-004","Ingreso","Venta Entradas - Fecha 26 vs Ramos Mejía Juv",28129,"2026-06-27","Efectivo"],["MOV-0312","T-004","Ingreso","Buffet/Cantina - Fecha 26 vs Ramos Mejía Juv",23285,"2026-06-27","Efectivo"],["MOV-0313","T-005","Gasto","Alquiler Cancha - Fecha 1 vs San Justo Baby",11962,"2026-01-03","Efectivo"],["MOV-0314","T-005","Gasto","Arbitraje - Fecha 1 vs San Justo Baby",6795,"2026-01-03","Efectivo"],["MOV-0315","T-005","Ingreso","Venta Entradas - Fecha 1 vs San Justo Baby",15080,"2026-01-03","Efectivo"],["MOV-0316","T-005","Ingreso","Buffet/Cantina - Fecha 1 vs San Justo Baby",13004,"2026-01-03","Efectivo"],["MOV-0317","T-005","Gasto","Traslado Delegación - Fecha 2 vs Ateneo Baby",6860,"2026-01-10","Efectivo"],["MOV-0318","T-005","Ingreso","Rifas y Buffet - Fecha 2 vs Ateneo Baby",9187,"2026-01-10","Efectivo"],["MOV-0319","T-005","Gasto","Alquiler Cancha - Fecha 3 vs Almafuerte Baby",11600,"2026-01-17","Efectivo"],["MOV-0320","T-005","Gasto","Arbitraje - Fecha 3 vs Almafuerte Baby",6646,"2026-01-17","Efectivo"],["MOV-0321","T-005","Ingreso","Venta Entradas - Fecha 3 vs Almafuerte Baby",16674,"2026-01-17","Efectivo"],["MOV-0322","T-005","Ingreso","Buffet/Cantina - Fecha 3 vs Almafuerte Baby",14764,"2026-01-17","Efectivo"],["MOV-0323","T-005","Gasto","Traslado Delegación - Fecha 4 vs Ituzaingó Baby",6772,"2026-01-24","Efectivo"],["MOV-0324","T-005","Ingreso","Rifas y Buffet - Fecha 4 vs Ituzaingó Baby",10016,"2026-01-24","Efectivo"],["MOV-0325","T-005","Gasto","Alquiler Cancha - Fecha 5 vs Liniers Baby",11506,"2026-01-31","Efectivo"],["MOV-0326","T-005","Gasto","Arbitraje - Fecha 5 vs Liniers Baby",7149,"2026-01-31","Efectivo"],["MOV-0327","T-005","Ingreso","Venta Entradas - Fecha 5 vs Liniers Baby",18247,"2026-01-31","Efectivo"],["MOV-0328","T-005","Ingreso","Buffet/Cantina - Fecha 5 vs Liniers Baby",15863,"2026-01-31","Efectivo"],["MOV-0329","T-005","Gasto","Traslado Delegación - Fecha 6 vs Almagro Baby",6409,"2026-02-07","Efectivo"],["MOV-0330","T-005","Ingreso","Rifas y Buffet - Fecha 6 vs Almagro Baby",10604,"2026-02-07","Efectivo"],["MOV-0331","T-005","Gasto","Alquiler Cancha - Fecha 7 vs Pearson Baby",11399,"2026-02-14","Efectivo"],["MOV-0332","T-005","Gasto","Arbitraje - Fecha 7 vs Pearson Baby",6686,"2026-02-14","Efectivo"],["MOV-0333","T-005","Ingreso","Venta Entradas - Fecha 7 vs Pearson Baby",19783,"2026-02-14","Efectivo"],["MOV-0334","T-005","Ingreso","Buffet/Cantina - Fecha 7 vs Pearson Baby",16490,"2026-02-14","Efectivo"],["MOV-0335","T-005","Gasto","Traslado Delegación - Fecha 8 vs CIDECO Baby",6811,"2026-02-21","Efectivo"],["MOV-0336","T-005","Ingreso","Rifas y Buffet - Fecha 8 vs CIDECO Baby",10780,"2026-02-21","Efectivo"],["MOV-0337","T-005","Gasto","Alquiler Cancha - Fecha 9 vs Estrella Baby",11210,"2026-02-28","Efectivo"],["MOV-0338","T-005","Gasto","Arbitraje - Fecha 9 vs Estrella Baby",6709,"2026-02-28","Efectivo"],["MOV-0339","T-005","Ingreso","Venta Entradas - Fecha 9 vs Estrella Baby",20796,"2026-02-28","Efectivo"],["MOV-0340","T-005","Ingreso","Buffet/Cantina - Fecha 9 vs Estrella Baby",16763,"2026-02-28","Efectivo"],["MOV-0341","T-005","Gasto","Traslado Delegación - Fecha 10 vs Defensores Baby",6457,"2026-03-07","Efectivo"],["MOV-0342","T-005","Ingreso","Rifas y Buffet - Fecha 10 vs Defensores Baby",12258,"2026-03-07","Efectivo"],["MOV-0343","T-005","Gasto","Alquiler Cancha - Fecha 11 vs Porteño Baby",12183,"2026-03-14","Efectivo"],["MOV-0344","T-005","Gasto","Arbitraje - Fecha 11 vs Porteño Baby",7043,"2026-03-14","Efectivo"],["MOV-0345","T-005","Ingreso","Venta Entradas - Fecha 11 vs Porteño Baby",21511,"2026-03-14","Efectivo"],["MOV-0346","T-005","Ingreso","Buffet/Cantina - Fecha 11 vs Porteño Baby",18886,"2026-03-14","Efectivo"],["MOV-0347","T-005","Gasto","Traslado Delegación - Fecha 12 vs Leloir Baby",6561,"2026-03-21","Efectivo"],["MOV-0348","T-005","Ingreso","Rifas y Buffet - Fecha 12 vs Leloir Baby",12223,"2026-03-21","Efectivo"],["MOV-0349","T-005","Gasto","Alquiler Cancha - Fecha 13 vs UAI Baby",12092,"2026-03-28","Efectivo"],["MOV-0350","T-005","Gasto","Arbitraje - Fecha 13 vs UAI Baby",6881,"2026-03-28","Efectivo"],["MOV-0351","T-005","Ingreso","Venta Entradas - Fecha 13 vs UAI Baby",22135,"2026-03-28","Efectivo"],["MOV-0352","T-005","Ingreso","Buffet/Cantina - Fecha 13 vs UAI Baby",18111,"2026-03-28","Efectivo"],["MOV-0353","T-005","Gasto","Traslado Delegación - Fecha 14 vs Pearson Juv Baby",6696,"2026-04-04","Efectivo"],["MOV-0354","T-005","Ingreso","Rifas y Buffet - Fecha 14 vs Pearson Juv Baby",12819,"2026-04-04","Efectivo"],["MOV-0355","T-005","Gasto","Alquiler Cancha - Fecha 15 vs Moron Futsal A Baby",11727,"2026-04-11","Efectivo"],["MOV-0356","T-005","Gasto","Arbitraje - Fecha 15 vs Moron Futsal A Baby",6836,"2026-04-11","Efectivo"],["MOV-0357","T-005","Ingreso","Venta Entradas - Fecha 15 vs Moron Futsal A Baby",23032,"2026-04-11","Efectivo"],["MOV-0358","T-005","Ingreso","Buffet/Cantina - Fecha 15 vs Moron Futsal A Baby",19104,"2026-04-11","Efectivo"],["MOV-0359","T-005","Gasto","Traslado Delegación - Fecha 16 vs Almafuerte Baby A Baby",6582,"2026-04-18","Efectivo"],["MOV-0360","T-005","Ingreso","Rifas y Buffet - Fecha 16 vs Almafuerte Baby A Baby",14339,"2026-04-18","Efectivo"],["MOV-0361","T-005","Gasto","Alquiler Cancha - Fecha 17 vs Lomas Futbol Baby",11503,"2026-04-25","Efectivo"],["MOV-0362","T-005","Gasto","Arbitraje - Fecha 17 vs Lomas Futbol Baby",6780,"2026-04-25","Efectivo"],["MOV-0363","T-005","Ingreso","Venta Entradas - Fecha 17 vs Lomas Futbol Baby",23264,"2026-04-25","Efectivo"],["MOV-0364","T-005","Ingreso","Buffet/Cantina - Fecha 17 vs Lomas Futbol Baby",19477,"2026-04-25","Efectivo"],["MOV-0365","T-005","Gasto","Traslado Delegación - Fecha 18 vs Porteño FutFem Baby",6717,"2026-05-02","Efectivo"],["MOV-0366","T-005","Ingreso","Rifas y Buffet - Fecha 18 vs Porteño FutFem Baby",14856,"2026-05-02","Efectivo"],["MOV-0367","T-005","Gasto","Alquiler Cancha - Fecha 19 vs Huracan San Justo Juv Baby",11381,"2026-05-09","Efectivo"],["MOV-0368","T-005","Gasto","Arbitraje - Fecha 19 vs Huracan San Justo Juv Baby",6812,"2026-05-09","Efectivo"],["MOV-0369","T-005","Ingreso","Venta Entradas - Fecha 19 vs Huracan San Justo Juv Baby",25137,"2026-05-09","Efectivo"],["MOV-0370","T-005","Ingreso","Buffet/Cantina - Fecha 19 vs Huracan San Justo Juv Baby",20343,"2026-05-09","Efectivo"],["MOV-0371","T-005","Gasto","Traslado Delegación - Fecha 20 vs Defensores Haedo Baby Baby",6853,"2026-05-16","Efectivo"],["MOV-0372","T-005","Ingreso","Rifas y Buffet - Fecha 20 vs Defensores Haedo Baby Baby",14761,"2026-05-16","Efectivo"],["MOV-0373","T-005","Gasto","Alquiler Cancha - Fecha 21 vs San Justo Baby Baby",11315,"2026-05-23","Efectivo"],["MOV-0374","T-005","Gasto","Arbitraje - Fecha 21 vs San Justo Baby Baby",6662,"2026-05-23","Efectivo"],["MOV-0375","T-005","Ingreso","Venta Entradas - Fecha 21 vs San Justo Baby Baby",26512,"2026-05-23","Efectivo"],["MOV-0376","T-005","Ingreso","Buffet/Cantina - Fecha 21 vs San Justo Baby Baby",22776,"2026-05-23","Efectivo"],["MOV-0377","T-005","Gasto","Traslado Delegación - Fecha 22 vs Huracán Baby",6839,"2026-05-30","Efectivo"],["MOV-0378","T-005","Ingreso","Rifas y Buffet - Fecha 22 vs Huracán Baby",15081,"2026-05-30","Efectivo"],["MOV-0379","T-005","Gasto","Alquiler Cancha - Fecha 23 vs Morón Baby",11530,"2026-06-06","Efectivo"],["MOV-0380","T-005","Gasto","Arbitraje - Fecha 23 vs Morón Baby",7196,"2026-06-06","Efectivo"],["MOV-0381","T-005","Ingreso","Venta Entradas - Fecha 23 vs Morón Baby",27261,"2026-06-06","Efectivo"],["MOV-0382","T-005","Ingreso","Buffet/Cantina - Fecha 23 vs Morón Baby",22983,"2026-06-06","Efectivo"],["MOV-0383","T-005","Gasto","Traslado Delegación - Fecha 24 vs Gimnasia Baby",6576,"2026-06-13","Efectivo"],["MOV-0384","T-005","Ingreso","Rifas y Buffet - Fecha 24 vs Gimnasia Baby",16676,"2026-06-13","Efectivo"],["MOV-0385","T-005","Gasto","Alquiler Cancha - Fecha 25 vs Ramos Mejía Baby",11582,"2026-06-20","Efectivo"],["MOV-0386","T-005","Gasto","Arbitraje - Fecha 25 vs Ramos Mejía Baby",6615,"2026-06-20","Efectivo"],["MOV-0387","T-005","Ingreso","Venta Entradas - Fecha 25 vs Ramos Mejía Baby",28682,"2026-06-20","Efectivo"],["MOV-0388","T-005","Ingreso","Buffet/Cantina - Fecha 25 vs Ramos Mejía Baby",23066,"2026-06-20","Efectivo"],["MOV-0389","T-005","Gasto","Traslado Delegación - Fecha 26 vs San Justo Baby",6617,"2026-06-27","Efectivo"],["MOV-0390","T-005","Ingreso","Rifas y Buffet - Fecha 26 vs San Justo Baby",16662,"2026-06-27","Efectivo"],["MOV-702746","General","Sponsor","publicidad",200000,"2026-06-27","Efectivo"]];
  finanzas.forEach(r => sheetFin.appendRow(r));
  
  // 7. Configurar Hoja "Partidos"
  let sheetPartidos = ss.getSheetByName(HOJA_PARTIDOS) || ss.insertSheet(HOJA_PARTIDOS);
  sheetPartidos.clear();
  sheetPartidos.appendRow(["Partido_ID","Torneo_ID","Date","Opponent","Location","Result","Scorers","Cards","MVP","Summary"]);
  const partidos = [["PAR-0001","T-001","2026-01-03","Huracán","Local","3 - 3","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol), Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol), Diaz, Lucas (Rival) (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a Huracán. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0002","T-001","2026-01-10","Morón","Visitante","4 - 4","Romero, Gonzalo (Rival) (2 goles), Torres, Mateo (Rival) (1 gol), Fernandez, Bruno (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Lopez, Javier (1 gol)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Morón. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0003","T-001","2026-01-17","Gimnasia","Local","3 - 0","Aguero, Lautaro (1 gol), Correa, Yamila (2 goles)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Gimnasia. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0004","T-001","2026-01-24","Ramos Mejía","Visitante","3 - 0","Fernandez, Lucas (Rival) (2 goles), Alvarez, Tomas (Rival) (1 gol)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a Ramos Mejía. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0005","T-001","2026-01-31","San Justo","Local","1 - 1","Aguero, Lautaro (1 gol), Benitez, Bruno (Rival) (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a San Justo. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0006","T-001","2026-02-07","Ateneo","Visitante","1 - 4","Sosa, Thiago (Rival) (1 gol), Aguero, Lautaro (2 goles), Correa, Yamila (2 goles)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Ateneo. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0007","T-001","2026-02-14","Almafuerte","Local","4 - 4","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (2 goles), Gomez, Pedro (Rival) (1 gol), Diaz, Julian (Rival) (2 goles), Sosa, Joaquin (Rival) (1 gol)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Almafuerte. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0008","T-001","2026-02-21","Ituzaingó","Visitante","5 - 3","Benitez, Bruno (Rival) (2 goles), Sanchez, Thiago (Rival) (1 gol), Romero, Nicolas (Rival) (2 goles), Aguero, Lautaro (2 goles), Correa, Yamila (1 gol)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Ituzaingó. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0009","T-001","2026-02-28","Liniers","Local","0 - 1","Ruiz, Julian (Rival) (1 gol)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Liniers. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0010","T-001","2026-03-07","Almagro","Visitante","3 - 4","Ruiz, Julian (Rival) (2 goles), Gomez, Joaquin (Rival) (1 gol), Aguero, Lautaro (2 goles), Correa, Yamila (2 goles)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Almagro. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0011","T-001","2026-03-14","Pearson","Local","1 - 1","Aguero, Lautaro (1 gol), Ruiz, Julian (Rival) (1 gol)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Pearson. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0012","T-001","2026-03-21","CIDECO","Visitante","2 - 1","Sosa, Thiago (Rival) (2 goles), Aguero, Lautaro (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a CIDECO. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0013","T-001","2026-03-28","Estrella","Local","4 - 0","Aguero, Lautaro (2 goles), Correa, Yamila (2 goles)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Estrella. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0014","T-001","2026-04-04","Defensores","Visitante","4 - 4","Gomez, Pedro (Rival) (1 gol), Diaz, Julian (Rival) (2 goles), Sosa, Joaquin (Rival) (1 gol), Aguero, Lautaro (2 goles), Correa, Yamila (2 goles)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Defensores. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0015","T-001","2026-04-11","Porteño","Local","0 - 2","Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a Porteño. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0016","T-001","2026-04-18","Leloir","Visitante","3 - 0","Sosa, Thiago (Rival) (1 gol), Benitez, Nicolas (Rival) (1 gol), Sanchez, Pedro (Rival) (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Leloir. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0017","T-001","2026-04-25","UAI","Local","1 - 2","Aguero, Lautaro (1 gol), Diaz, Nicolas (Rival) (1 gol), Sosa, Pedro (Rival) (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a UAI. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0018","T-001","2026-05-02","Pearson Juv","Visitante","3 - 0","Fernandez, Lucas (Rival) (2 goles), Alvarez, Tomas (Rival) (1 gol)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Pearson Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0019","T-001","2026-05-09","Moron Futsal A","Local","0 - 4","Alvarez, Joaquin (Rival) (1 gol), Ruiz, Lucas (Rival) (1 gol), Gomez, Tomas (Rival) (1 gol), Diaz, Gonzalo (Rival) (1 gol)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Moron Futsal A. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0020","T-001","2026-05-16","Almafuerte Baby A","Visitante","2 - 4","Ruiz, Julian (Rival) (2 goles), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (2 goles)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Almafuerte Baby A. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0021","T-001","2026-05-23","Lomas Futbol","Local","5 - 4","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (1 gol), Lopez, Alexis (1 gol), Messi, Sofia (1 gol), Sanchez, Mateo (Rival) (1 gol), Romero, Bruno (Rival) (1 gol), Torres, Thiago (Rival) (1 gol), Fernandez, Nicolas (Rival) (1 gol)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Lomas Futbol. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0022","T-001","2026-05-30","Porteño FutFem","Visitante","5 - 4","Alvarez, Joaquin (Rival) (2 goles), Ruiz, Lucas (Rival) (1 gol), Gomez, Tomas (Rival) (1 gol), Diaz, Gonzalo (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (2 goles)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Porteño FutFem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0023","T-001","2026-06-06","Huracan San Justo Juv","Local","Pendiente","","","",""],["PAR-0024","T-001","2026-06-13","Defensores Haedo Baby","Visitante","Pendiente","","","",""],["PAR-0025","T-001","2026-06-20","San Justo Baby","Local","Pendiente","","","",""],["PAR-0026","T-001","2026-06-27","Huracán","Visitante","Pendiente","","","",""],["PAR-0027","T-002","2026-01-03","Morón","Visitante","5 - 0","Romero, Gonzalo (Rival) (1 gol), Torres, Mateo (Rival) (1 gol), Fernandez, Bruno (Rival) (1 gol), Alvarez, Thiago (Rival) (2 goles)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Morón. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0028","T-002","2026-01-10","Gimnasia","Local","0 - 3","Torres, Tomas (Rival) (2 goles), Fernandez, Gonzalo (Rival) (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Gimnasia. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0029","T-002","2026-01-17","Ramos Mejía","Visitante","5 - 4","Fernandez, Lucas (Rival) (1 gol), Alvarez, Tomas (Rival) (1 gol), Ruiz, Gonzalo (Rival) (1 gol), Gomez, Mateo (Rival) (2 goles), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (2 goles)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Ramos Mejía. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0030","T-002","2026-01-24","San Justo","Local","4 - 1","Aguero, Lautaro (2 goles), Correa, Yamila (2 goles), Benitez, Bruno (Rival) (1 gol)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a San Justo. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0031","T-002","2026-01-31","Ateneo","Visitante","0 - 3","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (1 gol)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Ateneo. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0032","T-002","2026-02-07","Almafuerte","Local","2 - 4","Aguero, Lautaro (2 goles), Gomez, Pedro (Rival) (1 gol), Diaz, Julian (Rival) (2 goles), Sosa, Joaquin (Rival) (1 gol)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Almafuerte. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0033","T-002","2026-02-14","Ituzaingó","Visitante","3 - 4","Benitez, Bruno (Rival) (1 gol), Sanchez, Thiago (Rival) (1 gol), Romero, Nicolas (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Lopez, Javier (1 gol)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Ituzaingó. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0034","T-002","2026-02-21","Liniers","Local","3 - 0","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a Liniers. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0035","T-002","2026-02-28","Almagro","Visitante","5 - 3","Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol), Diaz, Lucas (Rival) (2 goles), Sosa, Tomas (Rival) (1 gol), Aguero, Lautaro (2 goles), Correa, Yamila (1 gol)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Almagro. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0036","T-002","2026-03-07","Pearson","Local","1 - 2","Aguero, Lautaro (1 gol), Ruiz, Julian (Rival) (2 goles)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Pearson. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0037","T-002","2026-03-14","CIDECO","Visitante","0 - 4","Aguero, Lautaro (2 goles), Correa, Yamila (2 goles)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a CIDECO. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0038","T-002","2026-03-21","Estrella","Local","4 - 0","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol), Lopez, Javier (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Estrella. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0039","T-002","2026-03-28","Defensores","Visitante","4 - 2","Gomez, Pedro (Rival) (2 goles), Diaz, Julian (Rival) (2 goles), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Defensores. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0040","T-002","2026-04-04","Porteño","Local","3 - 2","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (1 gol), Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Porteño. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0041","T-002","2026-04-11","Leloir","Visitante","1 - 2","Sosa, Thiago (Rival) (1 gol), Aguero, Lautaro (2 goles)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Leloir. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0042","T-002","2026-04-18","UAI","Local","0 - 2","Diaz, Nicolas (Rival) (2 goles)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a UAI. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0043","T-002","2026-04-25","Pearson Juv","Visitante","1 - 4","Fernandez, Lucas (Rival) (1 gol), Aguero, Lautaro (2 goles), Correa, Yamila (2 goles)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Pearson Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0044","T-002","2026-05-02","Moron Futsal A","Local","4 - 2","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol), Lopez, Javier (1 gol), Alvarez, Joaquin (Rival) (2 goles)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Moron Futsal A. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0045","T-002","2026-05-09","Almafuerte Baby A","Visitante","2 - 4","Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol), Aguero, Lautaro (2 goles), Correa, Yamila (1 gol), Lopez, Javier (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a Almafuerte Baby A. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0046","T-002","2026-05-16","Lomas Futbol","Local","2 - 4","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Sanchez, Mateo (Rival) (1 gol), Romero, Bruno (Rival) (1 gol), Torres, Thiago (Rival) (2 goles)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Lomas Futbol. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0047","T-002","2026-05-23","Porteño FutFem","Visitante","2 - 3","Alvarez, Joaquin (Rival) (2 goles), Aguero, Lautaro (2 goles), Correa, Yamila (1 gol)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Porteño FutFem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0048","T-002","2026-05-30","Huracan San Justo Juv","Local","4 - 4","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol), Lopez, Javier (1 gol), Fernandez, Lucas (Rival) (2 goles), Alvarez, Tomas (Rival) (2 goles)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a Huracan San Justo Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0049","T-002","2026-06-06","Defensores Haedo Baby","Visitante","Pendiente","","","",""],["PAR-0050","T-002","2026-06-13","San Justo Baby","Local","Pendiente","","","",""],["PAR-0051","T-002","2026-06-20","Huracán","Visitante","Pendiente","","","",""],["PAR-0052","T-002","2026-06-27","Morón","Local","Pendiente","","","",""],["PAR-0053","T-003","2026-01-03","Gimnasia Fem","Local","2 - 2","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Sanchez, Mateo (Rival) (1 gol), Romero, Bruno (Rival) (1 gol)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Gimnasia Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0054","T-003","2026-01-10","Ramos Mejía Fem","Visitante","1 - 1","Romero, Gonzalo (Rival) (1 gol), Aguero, Lautaro (1 gol)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Ramos Mejía Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0055","T-003","2026-01-17","San Justo Fem","Local","3 - 4","Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Diaz, Nicolas (Rival) (1 gol), Sosa, Pedro (Rival) (2 goles), Benitez, Julian (Rival) (1 gol)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a San Justo Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0056","T-003","2026-01-24","Ateneo Fem","Visitante","2 - 1","Gomez, Pedro (Rival) (2 goles), Aguero, Lautaro (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a Ateneo Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0057","T-003","2026-01-31","Almafuerte Fem","Local","3 - 3","Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Alvarez, Joaquin (Rival) (1 gol), Ruiz, Lucas (Rival) (1 gol), Gomez, Tomas (Rival) (1 gol)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Almafuerte Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0058","T-003","2026-02-07","Ituzaingó Fem","Visitante","0 - 1","Aguero, Lautaro (1 gol)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Ituzaingó Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0059","T-003","2026-02-14","Liniers Fem","Local","1 - 0","Aguero, Lautaro (1 gol)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a Liniers Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0060","T-003","2026-02-21","Almagro Fem","Visitante","2 - 3","Fernandez, Lucas (Rival) (2 goles), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Almagro Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0061","T-003","2026-02-28","Pearson Fem","Local","0 - 4","Fernandez, Lucas (Rival) (1 gol), Alvarez, Tomas (Rival) (1 gol), Ruiz, Gonzalo (Rival) (2 goles)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Pearson Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0062","T-003","2026-03-07","CIDECO Fem","Visitante","2 - 1","Gomez, Pedro (Rival) (2 goles), Aguero, Lautaro (1 gol)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a CIDECO Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0063","T-003","2026-03-14","Estrella Fem","Local","3 - 4","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol), Sanchez, Mateo (Rival) (2 goles), Romero, Bruno (Rival) (1 gol), Torres, Thiago (Rival) (1 gol)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Estrella Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0064","T-003","2026-03-21","Defensores Fem","Visitante","4 - 2","Alvarez, Joaquin (Rival) (1 gol), Ruiz, Lucas (Rival) (1 gol), Gomez, Tomas (Rival) (1 gol), Diaz, Gonzalo (Rival) (1 gol), Aguero, Lautaro (2 goles)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Defensores Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0065","T-003","2026-03-28","Porteño Fem","Local","1 - 3","Aguero, Lautaro (1 gol), Fernandez, Lucas (Rival) (2 goles), Alvarez, Tomas (Rival) (1 gol)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Porteño Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0066","T-003","2026-04-04","Leloir Fem","Visitante","5 - 3","Gomez, Pedro (Rival) (2 goles), Diaz, Julian (Rival) (1 gol), Sosa, Joaquin (Rival) (2 goles), Aguero, Lautaro (1 gol), Correa, Yamila (2 goles)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Leloir Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0067","T-003","2026-04-11","UAI Fem","Local","0 - 3","Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol), Diaz, Lucas (Rival) (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a UAI Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0068","T-003","2026-04-18","Pearson Juv Fem","Visitante","4 - 1","Romero, Gonzalo (Rival) (2 goles), Torres, Mateo (Rival) (2 goles), Aguero, Lautaro (1 gol)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Pearson Juv Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0069","T-003","2026-04-25","Moron Futsal A Fem","Local","2 - 1","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Torres, Tomas (Rival) (1 gol)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Moron Futsal A Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0070","T-003","2026-05-02","Almafuerte Baby A Fem","Visitante","3 - 1","Fernandez, Lucas (Rival) (2 goles), Alvarez, Tomas (Rival) (1 gol), Aguero, Lautaro (1 gol)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a Almafuerte Baby A Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0071","T-003","2026-05-09","Lomas Futbol Fem","Local","2 - 0","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Lomas Futbol Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0072","T-003","2026-05-16","Porteño FutFem Fem","Visitante","3 - 3","Torres, Tomas (Rival) (1 gol), Fernandez, Gonzalo (Rival) (2 goles), Aguero, Lautaro (2 goles), Correa, Yamila (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Porteño FutFem Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0073","T-003","2026-05-23","Huracan San Justo Juv Fem","Local","2 - 4","Aguero, Lautaro (2 goles), Romero, Gonzalo (Rival) (1 gol), Torres, Mateo (Rival) (2 goles), Fernandez, Bruno (Rival) (1 gol)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Huracan San Justo Juv Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0074","T-003","2026-05-30","Defensores Haedo Baby Fem","Visitante","5 - 0","Romero, Gonzalo (Rival) (1 gol), Torres, Mateo (Rival) (2 goles), Fernandez, Bruno (Rival) (2 goles)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Defensores Haedo Baby Fem. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0075","T-003","2026-06-06","San Justo Baby Fem","Local","Pendiente","","","",""],["PAR-0076","T-003","2026-06-13","Huracán Fem","Visitante","Pendiente","","","",""],["PAR-0077","T-003","2026-06-20","Morón Fem","Local","Pendiente","","","",""],["PAR-0078","T-003","2026-06-27","Gimnasia Fem","Visitante","Pendiente","","","",""],["PAR-0079","T-004","2026-01-03","Ramos Mejía Juv","Visitante","3 - 0","Romero, Gonzalo (Rival) (2 goles), Torres, Mateo (Rival) (1 gol)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Ramos Mejía Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0080","T-004","2026-01-10","San Justo Juv","Local","0 - 2","Diaz, Nicolas (Rival) (1 gol), Sosa, Pedro (Rival) (1 gol)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a San Justo Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0081","T-004","2026-01-17","Ateneo Juv","Visitante","5 - 4","Gomez, Pedro (Rival) (1 gol), Diaz, Julian (Rival) (2 goles), Sosa, Joaquin (Rival) (2 goles), Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Lopez, Javier (1 gol)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a Ateneo Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0082","T-004","2026-01-24","Almafuerte Juv","Local","3 - 4","Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Alvarez, Joaquin (Rival) (2 goles), Ruiz, Lucas (Rival) (2 goles)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Almafuerte Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0083","T-004","2026-01-31","Ituzaingó Juv","Visitante","2 - 0","Diaz, Nicolas (Rival) (1 gol), Sosa, Pedro (Rival) (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Ituzaingó Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0084","T-004","2026-02-07","Liniers Juv","Local","0 - 0","","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Liniers Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0085","T-004","2026-02-14","Almagro Juv","Visitante","3 - 3","Fernandez, Lucas (Rival) (1 gol), Alvarez, Tomas (Rival) (1 gol), Ruiz, Gonzalo (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (2 goles)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Almagro Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0086","T-004","2026-02-21","Pearson Juv","Local","4 - 2","Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Lopez, Javier (1 gol), Fernandez, Lucas (Rival) (1 gol), Alvarez, Tomas (Rival) (1 gol)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Pearson Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0087","T-004","2026-02-28","CIDECO Juv","Visitante","0 - 3","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a CIDECO Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0088","T-004","2026-03-07","Estrella Juv","Local","1 - 1","Aguero, Lautaro (1 gol), Sanchez, Mateo (Rival) (1 gol)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Estrella Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0089","T-004","2026-03-14","Defensores Juv","Visitante","4 - 2","Alvarez, Joaquin (Rival) (2 goles), Ruiz, Lucas (Rival) (1 gol), Gomez, Tomas (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a Defensores Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0090","T-004","2026-03-21","Porteño Juv","Local","0 - 0","","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Porteño Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0091","T-004","2026-03-28","Leloir Juv","Visitante","5 - 2","Gomez, Pedro (Rival) (2 goles), Diaz, Julian (Rival) (2 goles), Sosa, Joaquin (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Leloir Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0092","T-004","2026-04-04","UAI Juv","Local","0 - 4","Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol), Diaz, Lucas (Rival) (2 goles)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a UAI Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0093","T-004","2026-04-11","Pearson Juv Juv","Visitante","4 - 1","Romero, Gonzalo (Rival) (1 gol), Torres, Mateo (Rival) (2 goles), Fernandez, Bruno (Rival) (1 gol), Aguero, Lautaro (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Pearson Juv Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0094","T-004","2026-04-18","Moron Futsal A Juv","Local","2 - 3","Aguero, Lautaro (2 goles), Torres, Tomas (Rival) (1 gol), Fernandez, Gonzalo (Rival) (1 gol), Alvarez, Mateo (Rival) (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Moron Futsal A Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0095","T-004","2026-04-25","Almafuerte Baby A Juv","Visitante","1 - 0","Fernandez, Lucas (Rival) (1 gol)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Almafuerte Baby A Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0096","T-004","2026-05-02","Lomas Futbol Juv","Local","3 - 3","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (1 gol), Sosa, Thiago (Rival) (2 goles), Benitez, Nicolas (Rival) (1 gol)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Lomas Futbol Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0097","T-004","2026-05-09","Porteño FutFem Juv","Visitante","4 - 1","Torres, Tomas (Rival) (2 goles), Fernandez, Gonzalo (Rival) (2 goles), Aguero, Lautaro (1 gol)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Porteño FutFem Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0098","T-004","2026-05-16","Huracan San Justo Juv Juv","Local","1 - 4","Aguero, Lautaro (1 gol), Romero, Gonzalo (Rival) (1 gol), Torres, Mateo (Rival) (1 gol), Fernandez, Bruno (Rival) (2 goles)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Huracan San Justo Juv Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0099","T-004","2026-05-23","Defensores Haedo Baby Juv","Visitante","4 - 0","Romero, Gonzalo (Rival) (1 gol), Torres, Mateo (Rival) (1 gol), Fernandez, Bruno (Rival) (2 goles)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Defensores Haedo Baby Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0100","T-004","2026-05-30","San Justo Baby Juv","Local","4 - 3","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (1 gol), Lopez, Alexis (1 gol), Torres, Tomas (Rival) (1 gol), Fernandez, Gonzalo (Rival) (2 goles)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a San Justo Baby Juv. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0101","T-004","2026-06-06","Huracán Juv","Visitante","Pendiente","","","",""],["PAR-0102","T-004","2026-06-13","Morón Juv","Local","Pendiente","","","",""],["PAR-0103","T-004","2026-06-20","Gimnasia Juv","Visitante","Pendiente","","","",""],["PAR-0104","T-004","2026-06-27","Ramos Mejía Juv","Local","Pendiente","","","",""],["PAR-0105","T-005","2026-01-03","San Justo Baby","Local","0 - 3","Alvarez, Joaquin (Rival) (2 goles), Ruiz, Lucas (Rival) (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a San Justo Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0106","T-005","2026-01-10","Ateneo Baby","Visitante","0 - 2","Aguero, Lautaro (2 goles)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a Ateneo Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0107","T-005","2026-01-17","Almafuerte Baby","Local","2 - 2","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Romero, Gonzalo (Rival) (2 goles)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Almafuerte Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0108","T-005","2026-01-24","Ituzaingó Baby","Visitante","3 - 1","Alvarez, Joaquin (Rival) (1 gol), Ruiz, Lucas (Rival) (1 gol), Gomez, Tomas (Rival) (1 gol), Aguero, Lautaro (1 gol)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Ituzaingó Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0109","T-005","2026-01-31","Liniers Baby","Local","3 - 2","Aguero, Lautaro (2 goles), Correa, Yamila (1 gol), Sanchez, Mateo (Rival) (2 goles)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Liniers Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0110","T-005","2026-02-07","Almagro Baby","Visitante","5 - 1","Sanchez, Mateo (Rival) (1 gol), Romero, Bruno (Rival) (1 gol), Torres, Thiago (Rival) (1 gol), Fernandez, Nicolas (Rival) (2 goles), Aguero, Lautaro (1 gol)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Almagro Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0111","T-005","2026-02-14","Pearson Baby","Local","0 - 3","Sanchez, Mateo (Rival) (2 goles), Romero, Bruno (Rival) (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a Pearson Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0112","T-005","2026-02-21","CIDECO Baby","Visitante","2 - 0","Fernandez, Lucas (Rival) (1 gol), Alvarez, Tomas (Rival) (1 gol)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a CIDECO Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0113","T-005","2026-02-28","Estrella Baby","Local","1 - 0","Aguero, Lautaro (1 gol)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Estrella Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0114","T-005","2026-03-07","Defensores Baby","Visitante","5 - 3","Romero, Gonzalo (Rival) (1 gol), Torres, Mateo (Rival) (2 goles), Fernandez, Bruno (Rival) (2 goles), Aguero, Lautaro (2 goles), Correa, Yamila (1 gol)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a Defensores Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0115","T-005","2026-03-14","Porteño Baby","Local","1 - 3","Aguero, Lautaro (1 gol), Sanchez, Mateo (Rival) (2 goles), Romero, Bruno (Rival) (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Porteño Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0116","T-005","2026-03-21","Leloir Baby","Visitante","5 - 2","Fernandez, Lucas (Rival) (1 gol), Alvarez, Tomas (Rival) (1 gol), Ruiz, Gonzalo (Rival) (1 gol), Gomez, Mateo (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol)","Messi, Sofia 🟨 (Min 28)","Lopez, Alexis","Partido de alto ritmo competitivo frente a Leloir Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0117","T-005","2026-03-28","UAI Baby","Local","2 - 3","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Torres, Tomas (Rival) (1 gol), Fernandez, Gonzalo (Rival) (2 goles)","Lopez, Javier 🟨 (Min 28)","Rodriguez, Vanina","Partido de alto ritmo competitivo frente a UAI Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0118","T-005","2026-04-04","Pearson Juv Baby","Visitante","4 - 1","Sosa, Thiago (Rival) (1 gol), Benitez, Nicolas (Rival) (2 goles), Sanchez, Pedro (Rival) (1 gol), Aguero, Lautaro (1 gol)","Aguero, Lautaro 🟨 (Min 28)","Torres, Valentina","Partido de alto ritmo competitivo frente a Pearson Juv Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0119","T-005","2026-04-11","Moron Futsal A Baby","Local","5 - 2","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Lopez, Javier (1 gol), Lopez, Alexis (2 goles), Benitez, Bruno (Rival) (1 gol), Sanchez, Thiago (Rival) (1 gol)","Correa, Yamila 🟨 (Min 28)","Mac Allister, Joaquin","Partido de alto ritmo competitivo frente a Moron Futsal A Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0120","T-005","2026-04-18","Almafuerte Baby A Baby","Visitante","1 - 4","Sanchez, Mateo (Rival) (1 gol), Aguero, Lautaro (2 goles), Correa, Yamila (2 goles)","Lopez, Alexis 🟨 (Min 28)","Rodriguez, Carlos","Partido de alto ritmo competitivo frente a Almafuerte Baby A Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0121","T-005","2026-04-25","Lomas Futbol Baby","Local","2 - 3","Aguero, Lautaro (1 gol), Correa, Yamila (1 gol), Ruiz, Julian (Rival) (1 gol), Gomez, Joaquin (Rival) (1 gol), Diaz, Lucas (Rival) (1 gol)","Rodriguez, Vanina 🟨 (Min 28)","Zanetti, Milagros","Partido de alto ritmo competitivo frente a Lomas Futbol Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0122","T-005","2026-05-02","Porteño FutFem Baby","Visitante","4 - 1","Benitez, Bruno (Rival) (2 goles), Sanchez, Thiago (Rival) (1 gol), Romero, Nicolas (Rival) (1 gol), Aguero, Lautaro (1 gol)","Torres, Valentina 🟨 (Min 28)","Perez, Juan","Partido de alto ritmo competitivo frente a Porteño FutFem Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0123","T-005","2026-05-09","Huracan San Justo Juv Baby","Local","3 - 3","Aguero, Lautaro (1 gol), Correa, Yamila (2 goles), Sosa, Thiago (Rival) (2 goles), Benitez, Nicolas (Rival) (1 gol)","Mac Allister, Joaquin 🟨 (Min 28)","Messi, Sofia","Partido de alto ritmo competitivo frente a Huracan San Justo Juv Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0124","T-005","2026-05-16","Defensores Haedo Baby Baby","Visitante","3 - 1","Sosa, Thiago (Rival) (2 goles), Benitez, Nicolas (Rival) (1 gol), Aguero, Lautaro (1 gol)","Rodriguez, Carlos 🟨 (Min 28)","Lopez, Javier","Partido de alto ritmo competitivo frente a Defensores Haedo Baby Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0125","T-005","2026-05-23","San Justo Baby Baby","Local","4 - 3","Aguero, Lautaro (2 goles), Correa, Yamila (2 goles), Benitez, Bruno (Rival) (2 goles), Sanchez, Thiago (Rival) (1 gol)","Zanetti, Milagros 🟨 (Min 28)","Aguero, Lautaro","Partido de alto ritmo competitivo frente a San Justo Baby Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0126","T-005","2026-05-30","Huracán Baby","Visitante","3 - 2","Sanchez, Mateo (Rival) (2 goles), Romero, Bruno (Rival) (1 gol), Aguero, Lautaro (1 gol), Correa, Yamila (1 gol)","Perez, Juan 🟨 (Min 28)","Correa, Yamila","Partido de alto ritmo competitivo frente a Huracán Baby. El equipo demostró solidez táctica y gran contundencia ofensiva en momentos clave."],["PAR-0127","T-005","2026-06-06","Morón Baby","Local","Pendiente","","","",""],["PAR-0128","T-005","2026-06-13","Gimnasia Baby","Visitante","Pendiente","","","",""],["PAR-0129","T-005","2026-06-20","Ramos Mejía Baby","Local","Pendiente","","","",""],["PAR-0130","T-005","2026-06-27","San Justo Baby","Visitante","Pendiente","","","",""]];
  partidos.forEach(r => sheetPartidos.appendRow(r));

  return { success: true, message: "Base de datos inicializada con éxito con los 103 socios y datos completos." };
}

/**
 * Conecta a la base de datos Supabase mediante JDBC (usando IPv6 directa de Google Cloud)
 * para realizar la migración de la estructura (agregar username y password).
 */
function ejecutarMigracionSupabase() {
  const host = "db.kjcnotrxxthnzpgljeus.supabase.co";
  const port = 5432;
  const dbName = "postgres";
  const user = "postgres";
  const pass = "HaedoFutsal.2026";
  
  const url = "jdbc:postgresql://" + host + ":" + port + "/" + dbName;
  
  Logger.log("🔄 Conectando a Supabase mediante JDBC...");
  try {
    const conn = Jdbc.getConnection(url, user, pass);
    Logger.log("✅ Conexión JDBC establecida con éxito.");
    
    const stmt = conn.createStatement();
    
    Logger.log("🔄 Ejecutando ALTER TABLE...");
    stmt.execute(
      "ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS username TEXT UNIQUE;"
    );
    stmt.execute(
      "ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS password TEXT;"
    );
    Logger.log("✅ Columnas creadas con éxito.");
    
    Logger.log("🔄 Inicializando credenciales por defecto...");
    stmt.execute(
      "UPDATE usuarios SET username = split_part(email, '@', 1) WHERE username IS NULL;"
    );
    stmt.execute(
      "UPDATE usuarios SET password = '1234' WHERE password IS NULL;"
    );
    Logger.log("✅ Credenciales inicializadas.");
    
    stmt.close();
    conn.close();
    Logger.log("🎉 Migración de Supabase completada con éxito desde Apps Script!");
    return "✅ Migración ejecutada con éxito!";
  } catch (err) {
    Logger.log("❌ Error: " + err.message);
    throw new Error("Error en migración JDBC: " + err.message);
  }
}

/**
 * Actualiza los precios de cuotas de categorías y valor de entradas de torneos.
 * @param {Object} categoriasObj Mapa de Category_ID a Monthly_Fee
 * @param {Object} torneosObj Mapa de Torneo_ID a Ticket_Price
 */
function actualizarPrecios(categoriasObj, torneosObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Actualizar Categorías
    const sheetCategorias = ss.getSheetByName(HOJA_CATEGORIAS);
    if (sheetCategorias) {
      const data = sheetCategorias.getDataRange().getValues();
      const headers = data[0];
      const idIdx = headers.indexOf("Category_ID");
      const feeIdx = headers.indexOf("Monthly_Fee");
      
      if (idIdx !== -1 && feeIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          const catId = data[i][idIdx];
          if (categoriasObj[catId] !== undefined) {
            const newVal = parseFloat(categoriasObj[catId]) || 0;
            sheetCategorias.getRange(i + 1, feeIdx + 1).setValue(newVal);
          }
        }
      }
    }
    
    // 2. Actualizar Torneos
    const sheetTorneos = ss.getSheetByName(HOJA_TORNEOS);
    if (sheetTorneos) {
      const data = sheetTorneos.getDataRange().getValues();
      const headers = data[0];
      const idIdx = headers.indexOf("Torneo_ID");
      const nameIdx = headers.indexOf("Name");
      
      if (idIdx !== -1 && nameIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          const torneoId = data[i][idIdx];
          if (torneosObj[torneoId] !== undefined) {
            const rawName = data[i][nameIdx].toString();
            const nameLimpio = rawName.split(" [Precio:")[0].trim();
            const precio = parseFloat(torneosObj[torneoId]) || 0;
            const nuevoNombre = `${nameLimpio} [Precio: ${precio}]`;
            
            sheetTorneos.getRange(i + 1, nameIdx + 1).setValue(nuevoNombre);
          }
        }
      }
    }
    
    return { success: true, message: "Precios actualizados correctamente." };
  } catch (error) {
    console.error("Error en actualizarPrecios:", error);
    return { success: false, message: error.toString() };
  }
}
