/**
 * Google Apps Script – Seguimiento TMERT Plan 2026
 * Spreadsheet: 1cPeFZorUwiO3xXQmUwPhlV4Wy48Xg0xLOPV6xxoipBg
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 * 1. Abrir el spreadsheet en Google Sheets
 * 2. Extensions → Apps Script → pegar este código
 * 3. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copiar la URL generada y pegarla en secrets.toml como TMERT_API_URL
 */

const SPREADSHEET_ID  = '1cPeFZorUwiO3xXQmUwPhlV4Wy48Xg0xLOPV6xxoipBg';
const SHEET_SEGUIMIENTO = 'SeguimientoTMERT';   // Nombre exacto de la pestaña
const API_KEY           = 'tmert_seguimiento_IST2026';  // Cambiar si deseas otra clave

// ── GET ──────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    const key    = e.parameter.key;

    if (action === 'ping') {
      return json({ success: true, message: 'TMERT API OK', ts: new Date().toString() });
    }

    if (key !== API_KEY) {
      return json({ success: false, error: 'Clave API inválida' });
    }

    if (action === 'getInfo') {
      const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = getOrCreateSheet(ss);
      return json({
        success:  true,
        rows:     Math.max(0, sheet.getLastRow() - 1),
        updated:  sheet.getRange(2, 1).getValue() || 'sin datos'
      });
    }

    return json({ success: false, error: 'Acción no válida: ' + action });

  } catch (err) {
    return json({ success: false, error: err.toString() });
  }
}

// ── POST ─────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const action = e.parameter.action;
    const key    = e.parameter.key;

    if (key !== API_KEY) {
      return json({ success: false, error: 'Clave API inválida' });
    }

    if (!e.postData || !e.postData.contents) {
      return json({ success: false, error: 'Sin datos POST' });
    }

    const data = JSON.parse(e.postData.contents);

    switch (action) {
      case 'clearSheet':
        return clearSheet();

      case 'writeHeaders':
        return writeHeaders(data.headers);

      case 'appendRows':
        return appendRows(data.rows);

      case 'writeSeguimiento':
        // Escritura completa en una sola llamada (para datasets pequeños)
        return writeFull(data.headers, data.rows);

      default:
        return json({ success: false, error: 'Acción no válida: ' + action });
    }

  } catch (err) {
    return json({ success: false, error: err.toString(), stack: err.stack });
  }
}

// ── FUNCIONES INTERNAS ───────────────────────────────────────────────────────

function getOrCreateSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_SEGUIMIENTO);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SEGUIMIENTO);
  }
  return sheet;
}

function clearSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss);
  sheet.clearContents();
  return json({ success: true, message: 'Hoja limpiada' });
}

function writeHeaders(headers) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // Formato cabecera
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4F0B7B');
  headerRange.setFontColor('#FFFFFF');
  return json({ success: true, message: 'Cabecera escrita: ' + headers.length + ' columnas' });
}

function appendRows(rows) {
  if (!rows || rows.length === 0) {
    return json({ success: true, message: 'Sin filas para agregar' });
  }
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  return json({ success: true, rows_written: rows.length });
}

function writeFull(headers, rows) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss);
  sheet.clearContents();

  // Escribir cabecera
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4F0B7B');
  headerRange.setFontColor('#FFFFFF');

  // Escribir datos
  if (rows && rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  return json({
    success: true,
    message: 'Seguimiento actualizado',
    rows_written: rows ? rows.length : 0,
    timestamp: new Date().toString()
  });
}

// ── HELPER ───────────────────────────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
