// ============================================================
// DxD Backtesting Journal — Google Apps Script Backend
// วิธี Deploy:
//   1. เปิด Google Sheet → Extensions → Apps Script
//   2. วาง code นี้ทั้งหมด แทนที่ code เดิม
//   3. Deploy → New deployment → Web App
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy Web App URL ไปใส่ใน ⚙ Sheets ของ app
// ============================================================

const SHEET_NAME = 'trades';
const HEADERS    = [
  'id', 'datetime', 'asset', 'direction', 'session', 'setup',
  'entry', 'stop', 'exit', 'rr', 'status', 'beforeUrls', 'afterUrls'
];

// ============================================================
// ENTRY POINTS
// ============================================================

function doGet(e) {
  return respond(getAllTrades());
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { action, data } = payload;

    if (action === 'save')   return respond(saveTrade(data));
    if (action === 'update') return respond(updateTrade(data));
    if (action === 'delete') return respond(deleteTrade(data.id));

    return respond({ ok: false, error: `Unknown action: ${action}` });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

// ============================================================
// HELPER — JSON response
// ============================================================

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// HELPER — Get or create sheet with headers
// ============================================================

function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);

    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a1a2e');
    headerRange.setFontColor('#f5c542');
  }

  return sheet;
}

// ============================================================
// READ — Get all trades (newest first)
// ============================================================

function getAllTrades() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];   // header only

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  return rows
    .filter(row => row[0] !== '')  // skip empty rows
    .map(row => {
      const obj = {};
      HEADERS.forEach((h, i) => obj[h] = row[i]);

      // Parse JSON arrays
      try { obj.beforeUrls = JSON.parse(obj.beforeUrls || '[]'); } catch { obj.beforeUrls = []; }
      try { obj.afterUrls  = JSON.parse(obj.afterUrls  || '[]'); } catch { obj.afterUrls  = []; }

      // Cast numbers
      obj.id    = Number(obj.id);
      obj.entry = Number(obj.entry);
      obj.stop  = Number(obj.stop);
      obj.exit  = Number(obj.exit);
      obj.rr    = Number(obj.rr);

      return obj;
    });
}

// ============================================================
// CREATE — Append new trade row
// ============================================================

function saveTrade(trade) {
  if (!trade || !trade.id) return { ok: false, error: 'Missing trade data' };

  const sheet = getSheet();
  const row   = HEADERS.map(h => {
    if (h === 'beforeUrls' || h === 'afterUrls') return JSON.stringify(trade[h] || []);
    return trade[h] !== undefined ? trade[h] : '';
  });

  sheet.appendRow(row);
  return { ok: true, id: trade.id };
}

// ============================================================
// UPDATE — Find row by id and overwrite
// ============================================================

function updateTrade(trade) {
  if (!trade || !trade.id) return { ok: false, error: 'Missing trade id' };

  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { ok: false, error: 'No trades found' };

  const ids    = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
  const rowIdx = ids.indexOf(String(trade.id));

  if (rowIdx === -1) return { ok: false, error: `Trade ${trade.id} not found` };

  const sheetRow = rowIdx + 2;  // +2: 1-indexed + header row
  const rowData  = HEADERS.map(h => {
    if (h === 'beforeUrls' || h === 'afterUrls') return JSON.stringify(trade[h] || []);
    return trade[h] !== undefined ? trade[h] : '';
  });

  sheet.getRange(sheetRow, 1, 1, HEADERS.length).setValues([rowData]);
  return { ok: true, id: trade.id };
}

// ============================================================
// DELETE — Find row by id and remove
// ============================================================

function deleteTrade(id) {
  if (!id) return { ok: false, error: 'Missing id' };

  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { ok: false, error: 'No trades found' };

  const ids    = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
  const rowIdx = ids.indexOf(String(id));

  if (rowIdx === -1) return { ok: false, error: `Trade ${id} not found` };

  sheet.deleteRow(rowIdx + 2);
  return { ok: true, id };
}

// ============================================================
// UTILITY — Run this manually once to verify setup
// ============================================================

function testSetup() {
  const sheet = getSheet();
  Logger.log('Sheet name : ' + sheet.getName());
  Logger.log('Headers    : ' + sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0].join(', '));
  Logger.log('Trade rows : ' + Math.max(0, sheet.getLastRow() - 1));
  Logger.log('Setup OK ✓');
}
