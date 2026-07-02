// ============================================================
//  PORTAL CALIDAD - Google Apps Script API
//  Estructura de Google Sheets:
//    - Hoja "Usuarios"         : usuario | contraseña | rol | nombre
//    - Hoja "Guias"            : guías a procesar
//    - Hoja "BD_FUC"           : fichas únicas de campo
//    - Hoja "Programa_Proceso" : programa de proceso (fuente de estado)
// ============================================================

const SHEET_USUARIOS      = "Usuarios";
const SHEET_GUIAS         = "Guias";
const SHEET_FUC           = "BD_FUC";
const SHEET_PP            = "Programa_Proceso";
const SHEET_EVALUACIONES  = "BD_EVALUACIONES";
const SHEET_DEFECTOS      = "BD_DEFECTOS";
const SHEET_PRODUCTORES   = "BD_PRODUCTORES";
const SHEET_CALIBRES      = "BD_CALIBRES";
const SHEET_DESPACHOS     = "BD_DESPACHOS";

// ─────────────────────────────────────────────────────────────
//  ROUTER PRINCIPAL
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const body   = parseBody(e);
  const action = params.action || body.action || "";

  // Headers CORS
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    let result;
    switch (action) {
      case "login":          result = login(body);               break;
      case "getGuias":       result = getGuias(body);            break;
      case "addGuia":        result = addGuia(body);             break;
      case "updateGuia":     result = updateGuia(body);          break;
      case "deleteGuia":     result = deleteGuia(body);          break;
      case "uploadFUC":      result = uploadFUC(body);           break;
      case "getFUCs":        result = getFUCs(body);             break;
      case "deleteFUC":          result = deleteFUC(body);              break;
      case "getEvaluaciones":    result = getEvaluaciones(body);        break;
      case "addEvaluacion":      result = addEvaluacion(body);          break;
      case "updateEvaluacion":   result = updateEvaluacion(body);       break;
      case "deleteEvaluacion":   result = deleteEvaluacion(body);       break;
      case "getDefectos":        result = getDefectos(body);            break;
      case "addDefecto":         result = addDefecto(body);             break;
      case "updateDefecto":      result = updateDefecto(body);          break;
      case "deleteDefecto":        result = deleteDefecto(body);            break;
      case "getProductores":       result = getProductores(body);          break;
      case "uploadProductores":    result = uploadProductores(body);       break;
      case "deleteProductor":      result = deleteProductor(body);         break;
      case "getCalibres":          result = getCalibres(body);             break;
      case "addCalibre":           result = addCalibre(body);              break;
      case "updateCalibre":        result = updateCalibre(body);           break;
      case "deleteCalibre":        result = deleteCalibre(body);           break;
      case "getUsuarios":          result = getUsuarios(body);             break;
      case "addUsuario":           result = addUsuario(body);              break;
      case "updateUsuario":        result = updateUsuario(body);           break;
      case "deleteUsuario":        result = deleteUsuario(body);           break;
      case "getDespachos":         result = getDespachos(body);            break;
      case "addDespacho":          result = addDespacho(body);             break;
      case "updateDespacho":       result = updateDespacho(body);          break;
      case "deleteDespacho":       result = deleteDespacho(body);          break;
      default:
        result = { ok: false, error: "Acción no reconocida: " + action };
    }
    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return output;
}

// ─────────────────────────────────────────────────────────────
//  UTILIDADES
// ─────────────────────────────────────────────────────────────
function parseBody(e) {
  try {
    if (e.postData && e.postData.contents) {
      return JSON.parse(e.postData.contents);
    }
  } catch (_) {}
  return e.parameter || {};
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("Hoja '" + name + "' no encontrada en el spreadsheet.");
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => {
      const val = row[j];
      if (val instanceof Date) {
        if (val.getFullYear() <= 1899) {
          // Valor de solo hora (Sheets lo almacena como fecha 1899)
          // Formatear en hora local Perú para evitar desfase UTC
          obj[h] = Utilities.formatDate(val, "America/Lima", "HH:mm");
        } else {
          // Fecha/datetime completa: formatear en hora Perú
          obj[h] = Utilities.formatDate(val, "America/Lima", "yyyy-MM-dd'T'HH:mm:ss");
        }
      } else {
        obj[h] = val;
      }
    });
    return obj;
  });
}

function validateToken(token) {
  // Token simple: base64(usuario:timestamp) válido por 8 horas
  if (!token) throw new Error("Token requerido.");
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts   = decoded.split("|");
    if (parts.length < 2) throw new Error("Token inválido.");
    const ts = parseInt(parts[1]);
    if (Date.now() - ts > 8 * 60 * 60 * 1000) throw new Error("Sesión expirada.");
    return parts[0]; // usuario
  } catch (err) {
    throw new Error("Token inválido o expirado.");
  }
}

function generateToken(usuario) {
  const raw = usuario + "|" + Date.now();
  return Utilities.base64Encode(Utilities.newBlob(raw).getBytes());
}

// Retorna fecha/hora actual en zona horaria de Perú (America/Lima, UTC-5)
function nowPeru() {
  return Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd'T'HH:mm:ss");
}

// ─────────────────────────────────────────────────────────────
//  AUTENTICACIÓN
// ─────────────────────────────────────────────────────────────
function login(body) {
  const { usuario, password } = body;
  if (!usuario || !password) return { ok: false, error: "Usuario y contraseña requeridos." };

  const sheet = getSheet(SHEET_USUARIOS);
  const rows  = sheetToObjects(sheet);

  const user = rows.find(
    r => String(r.usuario).trim().toLowerCase() === String(usuario).trim().toLowerCase()
      && String(r.contraseña).trim() === String(password).trim()
  );

  if (!user) return { ok: false, error: "Credenciales incorrectas." };

  return {
    ok:      true,
    token:   generateToken(user.usuario),
    nombre:  user.nombre  || user.usuario,
    rol:     user.rol     || "operario",
    accesos: user.accesos || "all"   // "all" o lista separada por comas: "guias,eval,fuc"
  };
}

// ─────────────────────────────────────────────────────────────
//  MÓDULO GUÍAS
// ─────────────────────────────────────────────────────────────

// Columnas de la hoja "Guias" (en orden)
const GUIAS_HEADERS = [
  "VARIEDAD",
  "CLIENTE",
  "COD_PROD_ETIQ",   // antes FECHA_GUIA — ahora Cód. Prod. Etiq. de BD_PRODUCTORES
  "FECHA_INGRESO",
  "N_GUIA",
  "CLP",
  "PRODUCTOR",
  "LOTE",
  "CERT_LP",
  "MTD",
  "MATERIA_SECA",
  "QR_DUA",
  "VINCULACION_SISA",
  "PCT_DIF_QUEREME",
  "STATUS",
  "CREADO_POR",
  "FECHA_REGISTRO"
];

function initGuiasSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_GUIAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_GUIAS);
    sheet.appendRow(GUIAS_HEADERS);
    sheet.getRange(1, 1, 1, GUIAS_HEADERS.length).setFontWeight("bold");
  } else {
    // Auto-migración: actualizar cabecera si cambió (ej. FECHA_GUIA → COD_PROD_ETIQ)
    const lastCol    = Math.max(sheet.getLastColumn(), GUIAS_HEADERS.length);
    const headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const needsUpdate = GUIAS_HEADERS.some((h, i) => h !== headerVals[i]);
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, GUIAS_HEADERS.length).setValues([GUIAS_HEADERS]);
      sheet.getRange(1, 1, 1, GUIAS_HEADERS.length).setFontWeight("bold");
    }
  }
  return sheet;
}

function getGuias(body) {
  validateToken(body.token);
  const sheet = initGuiasSheet();
  const rows  = sheetToObjects(sheet);
  return { ok: true, data: rows };
}

function addGuia(body) {
  const usuario = validateToken(body.token);
  const sheet   = initGuiasSheet();
  const guia    = body.guia || {};

  const row = GUIAS_HEADERS.map(h => {
    if (h === "CREADO_POR")    return usuario;
    if (h === "FECHA_REGISTRO") return nowPeru();
    return guia[h] !== undefined ? guia[h] : "";
  });

  sheet.appendRow(row);
  return { ok: true, message: "Guía registrada correctamente." };
}

function updateGuia(body) {
  const usuario = validateToken(body.token);
  const sheet   = initGuiasSheet();
  const guia    = body.guia || {};
  const rowNum  = guia._row;

  if (!rowNum) return { ok: false, error: "_row requerido para actualizar." };

  const row = GUIAS_HEADERS.map(h => {
    if (h === "FECHA_REGISTRO") return guia[h] || nowPeru();
    return guia[h] !== undefined ? guia[h] : "";
  });

  sheet.getRange(rowNum, 1, 1, GUIAS_HEADERS.length).setValues([row]);
  return { ok: true, message: "Guía actualizada correctamente." };
}

function deleteGuia(body) {
  validateToken(body.token);
  const sheet  = initGuiasSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "Número de fila requerido." };
  sheet.deleteRow(rowNum);
  return { ok: true, message: "Guía eliminada." };
}

// ─────────────────────────────────────────────────────────────
//  MÓDULO BD_FUC
// ─────────────────────────────────────────────────────────────

const FUC_HEADERS = [
  "id", "certificado_organico", "codigo_lugar_produccion", "departamento",
  "empresa_certificadora", "exportador", "fecha_ingreso", "fecha_proceso",
  "fin_descarga", "ggn", "guia_ingreso", "guia_ingreso_local",
  "ingreso_garita", "inicio_descarga", "linea_empaque", "lote", "marca",
  "num_fuc", "num_personas", "pep", "producto", "productor",
  "tipo_caja", "trazabilidad", "variedad",
  "total_jabas", "kg_entregados", "kg_exportados",
  "estado", "created_at", "updated_at"
];

function initFUCSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_FUC);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FUC);
    sheet.appendRow(FUC_HEADERS);
    sheet.getRange(1, 1, 1, FUC_HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

// Obtener índice de columna en la hoja FUC (base 1)
function fucColIndex(key) {
  const idx = FUC_HEADERS.indexOf(key);
  return idx === -1 ? -1 : idx + 1;
}

// Buscar el estado en Programa_Proceso por num_fuc
function getEstadoFromPP(numFuc) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PP);
    if (!sheet) return "No programado";
    const rows = sheetToObjects(sheet);
    // Buscar columna que pueda llamarse num_fuc, fuc, n_fuc, etc.
    const match = rows.find(r => {
      const keys = Object.keys(r);
      return keys.some(k => {
        const kn = k.toLowerCase().replace(/[^a-z0-9]/g, "");
        return (kn === "numfuc" || kn === "fuc" || kn === "nfuc") &&
               String(r[k]).trim() === String(numFuc).trim();
      });
    });
    if (!match) return "No programado";
    // Buscar campo status_lote / estado / status
    const keys = Object.keys(match);
    const statusKey = keys.find(k => {
      const kn = k.toLowerCase().replace(/[^a-z0-9]/g, "");
      return kn === "statuslote" || kn === "estado" || kn === "status";
    });
    return statusKey ? (match[statusKey] || "No programado") : "No programado";
  } catch(_) {
    return "No programado";
  }
}

function getFUCs(body) {
  validateToken(body.token);
  const sheet = initFUCSheet();
  const rows  = sheetToObjects(sheet);
  return { ok: true, data: rows };
}

function uploadFUC(body) {
  validateToken(body.token);

  const rows = body.rows;
  if (!Array.isArray(rows) || rows.length === 0)
    return { ok: false, error: "Sin filas válidas para procesar." };

  const sheet    = initFUCSheet();
  const existing = sheetToObjects(sheet);

  // Índice por num_fuc para búsqueda rápida
  const byFUC = {};
  existing.forEach(r => {
    const k = String(r["num_fuc"] || "").trim();
    if (k) byFUC[k] = r;
  });

  let added = 0, updated = 0, skipped = 0;
  const now = nowPeru();

  rows.forEach(newRow => {
    const fuc = String(newRow["num_fuc"] || "").trim();
    if (!fuc) return;

    const old = byFUC[fuc];

    if (!old) {
      // INSERT
      const estado    = getEstadoFromPP(fuc);
      const id        = Utilities.getUuid();
      const rowValues = FUC_HEADERS.map(h => {
        if (h === "id")         return id;
        if (h === "estado")     return estado;
        if (h === "created_at") return now;
        if (h === "updated_at") return now;
        return newRow[h] !== undefined ? newRow[h] : "";
      });
      sheet.appendRow(rowValues);
      byFUC[fuc] = { _row: sheet.getLastRow(), ...newRow, id, estado, created_at: now, updated_at: now };
      added++;

    } else {
      // Comparar campos clave
      const changed =
        String(old["guia_ingreso"]  || "") !== String(newRow["guia_ingreso"]  || "") ||
        String(old["total_jabas"]   || "") !== String(newRow["total_jabas"]   || "") ||
        String(old["kg_entregados"] || "") !== String(newRow["kg_entregados"] || "");

      if (!changed) {
        skipped++;
      } else {
        // DELETE fila antigua + INSERT nueva actualizada
        sheet.deleteRow(old._row);

        // Recalcular filas existentes (los índices cambiaron)
        const refreshed = sheetToObjects(sheet);
        refreshed.forEach(r => { const k = String(r["num_fuc"]||"").trim(); if(k) byFUC[k]=r; });

        const estado    = getEstadoFromPP(fuc);
        const rowValues = FUC_HEADERS.map(h => {
          if (h === "id")         return old["id"] || Utilities.getUuid();
          if (h === "estado")     return estado;
          if (h === "created_at") return old["created_at"] || now;
          if (h === "updated_at") return now;
          return newRow[h] !== undefined ? newRow[h] : "";
        });
        sheet.appendRow(rowValues);
        updated++;
      }
    }
  });

  const total = added + updated + skipped;
  return {
    ok:      true,
    added,
    updated,
    skipped,
    total,
    message: `Procesadas ${total} FUCs: ${added} nuevas, ${updated} actualizadas, ${skipped} sin cambios.`
  };
}

function deleteFUC(body) {
  validateToken(body.token);
  const sheet  = initFUCSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "Número de fila requerido." };
  sheet.deleteRow(rowNum);
  return { ok: true, message: "FUC eliminada." };
}

// ─────────────────────────────────────────────────────────────
//  MÓDULO BD_EVALUACIONES
// ─────────────────────────────────────────────────────────────

const EVAL_HEADERS = [
  "id",
  "num_fuc",
  "num_evaluacion",   // número correlativo por FUC (1, 2, 3…)
  "exportador",
  "productor",
  "lote",
  "guia_ingreso",
  "cultivo",          // desde BD_FUC
  "variedad",         // desde BD_FUC
  "pct_exportable",   // % exportable salida de máquina 1 (Desviación de CAT)
  "pct_exportable_2", // % exportable salida de máquina 2 (Descarte)
  "pct_defectos",     // % defectos en PT
  "lineas",           // JSON: [{linea,linea_muestra,filas:[{calibre,cat,ps_defectos:[{cant,pct,def}],ss_defectos:[{cant,pct,def}]}]}]
  "observaciones",
  "created_by",
  "created_at",
  "updated_at",
  "maq1_muestra",     // cantidad de frutas muestreadas Máq. 1
  "maq1_encontradas", // cantidad encontradas (no exportable) Máq. 1
  "maq2_muestra",     // cantidad de frutas muestreadas Máq. 2
  "maq2_encontradas"  // cantidad encontradas (no exportable) Máq. 2
];

function initEvalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_EVALUACIONES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EVALUACIONES);
    sheet.appendRow(EVAL_HEADERS);
    sheet.getRange(1, 1, 1, EVAL_HEADERS.length).setFontWeight("bold");
  } else {
    // Auto-migración de cabeceras
    const lastCol    = Math.max(sheet.getLastColumn(), EVAL_HEADERS.length);
    const headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const needsUpdate = EVAL_HEADERS.some((h, i) => h !== headerVals[i]);
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, EVAL_HEADERS.length).setValues([EVAL_HEADERS]);
      sheet.getRange(1, 1, 1, EVAL_HEADERS.length).setFontWeight("bold");
    }
  }
  return sheet;
}

// Obtener todas las evaluaciones (opcionalmente filtradas por num_fuc)
function getEvaluaciones(body) {
  validateToken(body.token);
  const sheet = initEvalSheet();
  let rows    = sheetToObjects(sheet);
  if (body.num_fuc) {
    rows = rows.filter(r => String(r["num_fuc"]).trim() === String(body.num_fuc).trim());
  }
  return { ok: true, data: rows };
}

// Agregar nueva evaluación
function addEvaluacion(body) {
  const usuario = validateToken(body.token);
  const sheet   = initEvalSheet();
  const ev      = body.evaluacion || {};

  if (!ev.num_fuc) return { ok: false, error: "num_fuc es requerido." };

  // Calcular número correlativo para este FUC
  const all     = sheetToObjects(sheet);
  const prevNum = all
    .filter(r => String(r["num_fuc"]).trim() === String(ev.num_fuc).trim())
    .reduce((max, r) => Math.max(max, parseInt(r["num_evaluacion"]) || 0), 0);

  // lineas puede venir como array o string JSON
  const lineasStr = typeof ev.lineas === "string"
    ? ev.lineas
    : JSON.stringify(ev.lineas || []);

  const now = nowPeru();
  const row = EVAL_HEADERS.map(h => {
    if (h === "id")              return Utilities.getUuid();
    if (h === "num_evaluacion")  return prevNum + 1;
    if (h === "lineas")          return lineasStr;
    if (h === "created_by")      return usuario;
    if (h === "created_at")      return now;
    if (h === "updated_at")      return now;
    return ev[h] !== undefined ? ev[h] : "";
  });

  sheet.appendRow(row);
  return { ok: true, message: "Evaluación registrada.", num_evaluacion: prevNum + 1 };
}

// Actualizar evaluación existente
function updateEvaluacion(body) {
  validateToken(body.token);
  const sheet  = initEvalSheet();
  const ev     = body.evaluacion || {};
  const rowNum = ev._row;

  if (!rowNum) return { ok: false, error: "_row requerido." };

  // Leer fila existente para preservar created_at original
  const existing = {};
  const existingVals = sheet.getRange(rowNum, 1, 1, EVAL_HEADERS.length).getValues()[0];
  EVAL_HEADERS.forEach(function(h, i) { existing[h] = existingVals[i]; });

  const lineasStr = typeof ev.lineas === "string"
    ? ev.lineas
    : JSON.stringify(ev.lineas || []);

  const now = nowPeru();
  const row = EVAL_HEADERS.map(h => {
    if (h === "updated_at") return now;
    if (h === "created_at") return existing["created_at"] || now; // nunca pisar la fecha original
    if (h === "lineas")     return lineasStr;
    return ev[h] !== undefined ? ev[h] : "";
  });

  sheet.getRange(rowNum, 1, 1, EVAL_HEADERS.length).setValues([row]);
  return { ok: true, message: "Evaluación actualizada." };
}

// Eliminar evaluación
function deleteEvaluacion(body) {
  validateToken(body.token);
  const sheet  = initEvalSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "Número de fila requerido." };
  sheet.deleteRow(rowNum);
  return { ok: true, message: "Evaluación eliminada." };
}

// ─────────────────────────────────────────────────────────────
//  MIGRACIÓN BD_EVALUACIONES (ejecutar una vez desde el editor)
// ─────────────────────────────────────────────────────────────
// Cabeceras OLD (14 cols) que pueden existir en el sheet
const OLD_EVAL_HEADERS = [
  "id","num_fuc","num_evaluacion","exportador","productor","lote","guia_ingreso",
  "pct_exportable","pct_defectos","defectos","observaciones",
  "created_by","created_at","updated_at"
];

function migrateEvalSheet() {
  const sheet = initEvalSheet(); // actualiza header si es necesario
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("Sin filas que migrar."); return; }

  let migrated = 0;
  for (let r = 2; r <= lastRow; r++) {
    const rowData = sheet.getRange(r, 1, 1, Math.max(sheet.getLastColumn(), 16)).getValues()[0];

    // Detectar fila OLD: si col 8 (idx 7, cultivo) contiene un número → es pct_exportable viejo
    const maybeCultivo = String(rowData[7] || "").trim();
    const isOldFormat  = maybeCultivo !== "" && !isNaN(parseFloat(maybeCultivo));

    if (!isOldFormat) continue; // ya está en nuevo formato

    // Leer valores en posiciones OLD (0-indexed)
    const old_pct_exportable = rowData[7];  // old col 8
    const old_pct_defectos   = rowData[8];  // old col 9
    const old_defectos       = rowData[9];  // old col 10
    const old_observaciones  = rowData[10]; // old col 11
    const old_created_by     = rowData[11]; // old col 12
    const old_created_at     = rowData[12]; // old col 13
    const old_updated_at     = rowData[13]; // old col 14

    // Construir nueva fila (16 cols)
    const newRow = [
      rowData[0],          // id
      rowData[1],          // num_fuc
      rowData[2],          // num_evaluacion
      rowData[3],          // exportador
      rowData[4],          // productor
      rowData[5],          // lote
      rowData[6],          // guia_ingreso
      "",                  // cultivo (desconocido para filas viejas)
      "",                  // variedad (desconocido)
      old_pct_exportable,  // pct_exportable
      old_pct_defectos,    // pct_defectos (se mantiene)
      JSON.stringify([{linea:1, observaciones: String(old_defectos||""), filas:[]}]), // lineas: defectos viejos como obs
      old_observaciones,   // observaciones
      old_created_by,      // created_by
      old_created_at,      // created_at
      old_updated_at,      // updated_at
    ];

    sheet.getRange(r, 1, 1, 16).setValues([newRow]);
    migrated++;
  }

  SpreadsheetApp.getUi().alert(`✅ Migración completada: ${migrated} fila(s) actualizadas.`);
}

// ─────────────────────────────────────────────────────────────
//  MÓDULO BD_PRODUCTORES
// ─────────────────────────────────────────────────────────────

const PRODUCTORES_HDR = [
  "id", "variedad", "equiv_productor", "cod_productor_etiqueta",
  "cod_lote", "cliente", "clp", "productor", "lote",
  "created_at", "updated_at"
];

function initProductoresSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PRODUCTORES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PRODUCTORES);
    sheet.appendRow(PRODUCTORES_HDR);
    sheet.getRange(1, 1, 1, PRODUCTORES_HDR.length).setFontWeight("bold");
  }
  return sheet;
}

function getProductores(body) {
  validateToken(body.token);
  const sheet = initProductoresSheet();
  return { ok: true, data: sheetToObjects(sheet) };
}

function uploadProductores(body) {
  validateToken(body.token);
  const rows = body.rows;
  if (!Array.isArray(rows) || !rows.length)
    return { ok: false, error: "Sin filas válidas para procesar." };

  const sheet = initProductoresSheet();
  let added   = 0;
  const now   = nowPeru();

  // Sin verificación de duplicados — se cargan todos los registros con CLP válido
  rows.forEach(newRow => {
    const clp = String(newRow["clp"] || "").trim();
    if (!clp) return;   // fila sin CLP → ignorar
    const rowValues = PRODUCTORES_HDR.map(h => {
      if (h === "id")         return Utilities.getUuid();
      if (h === "created_at") return now;
      if (h === "updated_at") return now;
      return newRow[h] !== undefined ? newRow[h] : "";
    });
    sheet.appendRow(rowValues);
    added++;
  });

  return { ok: true, added, skipped: 0, duplicates: [], total: added,
           message: `${added} registros añadidos correctamente.` };
}

function deleteProductor(body) {
  validateToken(body.token);
  const sheet  = initProductoresSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "Número de fila requerido." };
  sheet.deleteRow(rowNum);
  return { ok: true, message: "Registro eliminado." };
}

// ─────────────────────────────────────────────────────────────
//  MÓDULO BD_DEFECTOS
// ─────────────────────────────────────────────────────────────

const DEFECTOS_HDR = ["id", "cultivo", "defecto", "activo", "created_by", "created_at", "updated_at"];

function initDefectosSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DEFECTOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DEFECTOS);
    sheet.appendRow(DEFECTOS_HDR);
    sheet.getRange(1, 1, 1, DEFECTOS_HDR.length).setFontWeight("bold");
  }
  return sheet;
}

// Obtener defectos; opcionalmente filtrar por cultivo
function getDefectos(body) {
  validateToken(body.token);
  const sheet = initDefectosSheet();
  let rows    = sheetToObjects(sheet);
  if (body.cultivo) {
    const c = String(body.cultivo).trim().toUpperCase();
    rows = rows.filter(r => String(r["cultivo"]).trim().toUpperCase() === c);
  }
  // Solo activos por defecto; pasar all=true para traer todos
  if (!body.all) {
    rows = rows.filter(r => String(r["activo"]).trim().toUpperCase() !== "FALSE" &&
                            String(r["activo"]).trim() !== "0");
  }
  return { ok: true, data: rows };
}

function addDefecto(body) {
  const usuario = validateToken(body.token);
  const sheet   = initDefectosSheet();
  const d       = body.defecto || {};
  if (!d.cultivo || !d.defecto)
    return { ok: false, error: "cultivo y defecto son requeridos." };

  const now = nowPeru();
  const row = DEFECTOS_HDR.map(h => {
    if (h === "id")         return Utilities.getUuid();
    if (h === "activo")     return true;
    if (h === "created_by") return usuario;
    if (h === "created_at") return now;
    if (h === "updated_at") return now;
    return d[h] !== undefined ? d[h] : "";
  });
  sheet.appendRow(row);
  return { ok: true, message: "Defecto agregado." };
}

function updateDefecto(body) {
  validateToken(body.token);
  const sheet  = initDefectosSheet();
  const d      = body.defecto || {};
  const rowNum = d._row;
  if (!rowNum) return { ok: false, error: "_row requerido." };
  const now = nowPeru();
  const row = DEFECTOS_HDR.map(h => {
    if (h === "updated_at") return now;
    return d[h] !== undefined ? d[h] : "";
  });
  sheet.getRange(rowNum, 1, 1, DEFECTOS_HDR.length).setValues([row]);
  return { ok: true, message: "Defecto actualizado." };
}

function deleteDefecto(body) {
  validateToken(body.token);
  const sheet  = initDefectosSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "Número de fila requerido." };
  sheet.deleteRow(rowNum);
  return { ok: true, message: "Defecto eliminado." };
}

// ─────────────────────────────────────────────────────────────
//  INICIALIZACIÓN (ejecutar una sola vez manualmente)
// ─────────────────────────────────────────────────────────────
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hoja Usuarios
  let usuarios = ss.getSheetByName(SHEET_USUARIOS);
  if (!usuarios) {
    usuarios = ss.insertSheet(SHEET_USUARIOS);
    usuarios.appendRow(["usuario", "contraseña", "rol", "nombre"]);
    usuarios.appendRow(["admin", "admin123", "admin", "Administrador"]);
    usuarios.getRange(1, 1, 1, 4).setFontWeight("bold");
  }

  // Hoja Guias
  initGuiasSheet();

  // Hoja BD_FUC
  initFUCSheet();

  // Hoja BD_EVALUACIONES
  initEvalSheet();

  // Hoja BD_DEFECTOS
  initDefectosSheet();

  // Hoja BD_PRODUCTORES
  initProductoresSheet();

  // Hoja BD_CALIBRES
  initCalibresSheet();

  SpreadsheetApp.getUi().alert("✅ Estructura creada correctamente.");
}

// ─────────────────────────────────────────────────────────────
//  MÓDULO BD_CALIBRES
// ─────────────────────────────────────────────────────────────

const CALIBRES_HEADERS = ["cultivo", "calibre", "descripcion", "created_at"];

function initCalibresSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CALIBRES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CALIBRES);
    sheet.appendRow(CALIBRES_HEADERS);
    sheet.getRange(1, 1, 1, CALIBRES_HEADERS.length).setFontWeight("bold");
  } else {
    // Auto-migración: agregar columna "cultivo" si no existe
    const h = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CALIBRES_HEADERS.length)).getValues()[0];
    if (CALIBRES_HEADERS.some((v, i) => v !== h[i])) {
      sheet.getRange(1, 1, 1, CALIBRES_HEADERS.length).setValues([CALIBRES_HEADERS]);
      sheet.getRange(1, 1, 1, CALIBRES_HEADERS.length).setFontWeight("bold");
    }
  }
  return sheet;
}

function getCalibres(body) {
  validateToken(body.token);
  const sheet = initCalibresSheet();
  let rows = sheetToObjects(sheet);
  if (body.cultivo) {
    const c = String(body.cultivo).trim().toUpperCase();
    rows = rows.filter(r => String(r.cultivo || "").trim().toUpperCase() === c);
  }
  return { ok: true, data: rows };
}

function addCalibre(body) {
  validateToken(body.token);
  const sheet   = initCalibresSheet();
  const cal     = String(body.calibre || "").trim();
  const cultivo = String(body.cultivo  || "").trim().toUpperCase();
  if (!cal)     return { ok: false, error: "Calibre es requerido." };
  if (!cultivo) return { ok: false, error: "Cultivo es requerido." };

  const all = sheetToObjects(sheet);
  if (all.some(r =>
    String(r.calibre||"").trim().toUpperCase() === cal.toUpperCase() &&
    String(r.cultivo||"").trim().toUpperCase() === cultivo
  )) return { ok: false, error: `El calibre "${cal}" ya existe para ${cultivo}.` };

  const now = nowPeru();
  sheet.appendRow([cultivo, cal, body.descripcion || "", now]);
  return { ok: true, message: "Calibre agregado." };
}

function updateCalibre(body) {
  validateToken(body.token);
  const sheet   = initCalibresSheet();
  const rowNum  = body.row;
  if (!rowNum) return { ok: false, error: "_row requerido." };
  const cal     = String(body.calibre || "").trim();
  const cultivo = String(body.cultivo  || "").trim().toUpperCase();
  if (!cal) return { ok: false, error: "Calibre es requerido." };
  const now = nowPeru();
  sheet.getRange(rowNum, 1, 1, CALIBRES_HEADERS.length)
    .setValues([[cultivo, cal, body.descripcion || "", now]]);
  return { ok: true, message: "Calibre actualizado." };
}

function deleteCalibre(body) {
  validateToken(body.token);
  const sheet  = initCalibresSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "_row requerido." };
  sheet.deleteRow(rowNum);
  return { ok: true, message: "Calibre eliminado." };
}

// ─────────────────────────────────────────────────────────────
//  MÓDULO GESTIÓN DE USUARIOS
// ─────────────────────────────────────────────────────────────

const USUARIOS_HDR_EXT = ["usuario", "contraseña", "rol", "nombre", "accesos"];

function initUsuariosSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USUARIOS);
  if (!sheet) throw new Error("Hoja Usuarios no encontrada.");
  // Agregar columna "accesos" si no existe
  const lastCol = sheet.getLastColumn();
  if (lastCol < 5) {
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headers.includes("accesos")) {
      sheet.getRange(1, 5).setValue("accesos");
      sheet.getRange(1, 5).setFontWeight("bold");
    }
  }
  return sheet;
}

function getUsuarios(body) {
  validateToken(body.token);
  const sheet = initUsuariosSheet();
  const rows  = sheetToObjects(sheet);
  // No devolver contraseñas
  return { ok: true, data: rows.map(r => ({
    _row:    r._row,
    usuario: r.usuario,
    rol:     r.rol,
    nombre:  r.nombre,
    accesos: r.accesos || "all"
  }))};
}

function addUsuario(body) {
  validateToken(body.token);
  const sheet = initUsuariosSheet();
  const u     = String(body.usuario   || "").trim().toLowerCase();
  const pwd   = String(body.contrasena|| "").trim();
  const rol   = String(body.rol       || "usuario").trim();
  const nombre= String(body.nombre    || "").trim();
  const acc   = String(body.accesos   || "all").trim();
  if (!u)   return { ok: false, error: "Usuario es requerido." };
  if (!pwd) return { ok: false, error: "Contraseña es requerida." };
  const all = sheetToObjects(sheet);
  if (all.some(r => String(r.usuario||"").trim().toLowerCase() === u))
    return { ok: false, error: `El usuario "${u}" ya existe.` };
  sheet.appendRow([u, pwd, rol, nombre, acc]);
  return { ok: true, message: "Usuario creado." };
}

function updateUsuario(body) {
  validateToken(body.token);
  const sheet  = initUsuariosSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "_row requerido." };
  const existing = sheetToObjects(sheet).find(r => r._row === rowNum);
  const pwd = body.contrasena ? String(body.contrasena).trim() : (existing ? existing["contraseña"] : "");
  sheet.getRange(rowNum, 1, 1, 5).setValues([[
    String(body.usuario  || "").trim().toLowerCase(),
    pwd,
    String(body.rol      || "usuario").trim(),
    String(body.nombre   || "").trim(),
    String(body.accesos  || "all").trim()
  ]]);
  return { ok: true, message: "Usuario actualizado." };
}

function deleteUsuario(body) {
  validateToken(body.token);
  const sheet  = initUsuariosSheet();
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "_row requerido." };
  // No permitir borrar la única cuenta admin
  const rows = sheetToObjects(sheet);
  const admins = rows.filter(r => r.rol === "admin");
  const target = rows.find(r => r._row === rowNum);
  if (target && target.rol === "admin" && admins.length <= 1)
    return { ok: false, error: "No puedes eliminar el único administrador." };
  sheet.deleteRow(rowNum);
  return { ok: true, message: "Usuario eliminado." };
}

// ═══════════════════════════════════════════════════════════════
//  MÓDULO BD_DESPACHOS
// ═══════════════════════════════════════════════════════════════

const DESPACHOS_HEADERS = [
  "id", "fecha", "num_contenedor", "num_booking", "exportador",
  "hora_programada", "hora_entrega_carga", "hora_inicio_revision_calidad",
  "hora_inicio_correccion", "hora_fin_correccion", "areas_involucradas",
  "hora_entrega_anexo", "hora_validacion_anexo",
  "hora_inicio_revision_senasa", "hora_fin_revision_senasa",
  "observaciones", "created_by", "created_at"
];

function initDespachoSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_DESPACHOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DESPACHOS);
    sheet.appendRow(DESPACHOS_HEADERS);
    sheet.getRange(1, 1, 1, DESPACHOS_HEADERS.length)
         .setBackground("#1b5e20").setFontColor("#ffffff").setFontWeight("bold");
    return sheet;
  }
  // Auto-migración si faltan columnas
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                        .map(h => String(h).trim().toLowerCase());
  DESPACHOS_HEADERS.forEach((h, i) => {
    const col = i + 1;
    const idx = existing.indexOf(h.toLowerCase());
    if (idx === -1) {
      if (col > sheet.getLastColumn()) sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, col).setValue(h);
    }
  });
  return sheet;
}

function getDespachos(body) {
  validateToken(body.token);
  const sheet = initDespachoSheet();
  const rows  = sheetToObjects(sheet);
  return { ok: true, rows: rows };
}

function addDespacho(body) {
  const usuario = validateToken(body.token);
  const d   = body;
  const now = nowPeru();
  const sheet = initDespachoSheet();
  const rowValues = DESPACHOS_HEADERS.map(h => {
    if (h === "id")         return Utilities.getUuid();
    if (h === "created_by") return usuario;
    if (h === "created_at") return now;
    return d[h] !== undefined ? d[h] : "";
  });
  sheet.appendRow(rowValues);
  return { ok: true, message: "Despacho registrado." };
}

function updateDespacho(body) {
  validateToken(body.token);
  const targetId = body.id;
  if (!targetId) return { ok: false, error: "id requerido." };
  const sheet = initDespachoSheet();
  const rows  = sheetToObjects(sheet);
  const orig  = rows.find(r => r.id === targetId);
  if (!orig) return { ok: false, error: "Despacho no encontrado." };
  const rowNum = orig._row;
  const rowValues = DESPACHOS_HEADERS.map(h => {
    if (h === "id")         return orig.id;
    if (h === "created_at") return orig.created_at || nowPeru();
    if (h === "created_by") return orig.created_by || "";
    return body[h] !== undefined ? body[h] : (orig[h] || "");
  });
  sheet.getRange(rowNum, 1, 1, DESPACHOS_HEADERS.length).setValues([rowValues]);
  return { ok: true, message: "Despacho actualizado." };
}

function deleteDespacho(body) {
  validateToken(body.token);
  const rowNum = body.row;
  if (!rowNum) return { ok: false, error: "row requerido." };
  const sheet = initDespachoSheet();
  sheet.deleteRow(parseInt(rowNum));
  return { ok: true, message: "Despacho eliminado." };
}

// ─────────────────────────────────────────────────────────────
//  UTILIDAD: corregir created_at de evaluaciones específicas
//  Ejecutar UNA SOLA VEZ desde el editor de GAS y luego borrar
// ─────────────────────────────────────────────────────────────
function fixCreatedAtEvaluaciones() {
  const FUCS_A_CORREGIR = ["WM-000032", "WM-000031", "TA-000024", "WM-000009"];
  const FECHA_CORRECTA  = "2026-06-28T00:00:00";

  const sheet = initEvalSheet();
  const data  = sheet.getDataRange().getValues();
  const hdrs  = data[0];
  const colFuc = hdrs.indexOf("num_fuc");
  const colCat = hdrs.indexOf("created_at");

  if (colFuc === -1 || colCat === -1) {
    Logger.log("No se encontraron las columnas num_fuc o created_at.");
    return;
  }

  let actualizados = 0;
  for (let i = 1; i < data.length; i++) {
    const fuc = String(data[i][colFuc] || "").trim();
    if (FUCS_A_CORREGIR.includes(fuc)) {
      sheet.getRange(i + 1, colCat + 1).setValue(FECHA_CORRECTA);
      Logger.log("Actualizado: fila " + (i + 1) + " — " + fuc);
      actualizados++;
    }
  }
  Logger.log("Total actualizados: " + actualizados);
}
