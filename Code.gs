// BLUE RAVEN AG — Google Apps Script v2.0
// Paste into script.google.com — deploy as Web App
//   Execute as: Me
//   Who has access: Anyone (anonymous)
// After ANY code change: Deploy → Manage Deployments → Edit → New Version → Deploy

const SHEET_ID = '1txSQJVw7nA42BtS8ilDyrM4H0nYtGWvP3mb4tcZJ034';

const TABS = {
  orders:        'Orders',
  customers:     'Customers',
  pilots:        'Pilots',
  products:      'Products',
  templates:     'MixTemplates',
  templateProds: 'MixTemplateProducts',
  orderProds:    'OrderProducts',
  fields:        'Fields',
  orderFields:   'OrderFields',
};

// ── doGet: reads all data OR handles small writes via query params ─────────
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'read';
    if (action === 'read')  return handleRead();
    if (action === 'write') return handleWrite(e.parameter);
    return jsonOut({ error: 'Unknown action: ' + action });
  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

// ── doPost: handles large writes (e.g. polygon KML data) ──────────────────
// Called with Content-Type: text/plain, body = JSON string
function doPost(e) {
  try {
    const raw     = e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    const { action, table, data } = payload;
    if (action === 'write') {
      return handleWriteData(table, data);
    }
    return jsonOut({ error: 'Unknown POST action: ' + action });
  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

function handleRead() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const result = {};
  Object.entries(TABS).forEach(([key, tabName]) => {
    const sheet = ss.getSheetByName(tabName);
    result[key] = sheet ? sheet.getDataRange().getValues() : [];
  });
  return jsonOut(result);
}

function handleWrite(params) {
  const table = params.table;
  const data  = JSON.parse(decodeURIComponent(params.data));
  return handleWriteData(table, data);
}

function handleWriteData(table, data) {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const tabName = TABS[table];
  const sheet   = ss.getSheetByName(tabName);
  if (!sheet) return jsonOut({ error: 'Sheet not found: ' + tabName });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();

  // ── Special upsert for templateProds: match on TemplateID + ProductID ──
  if (table === 'templateProds' && !data._delete) {
    const tmplCol = headers.indexOf('TemplateID');
    const prodCol = headers.indexOf('ProductID');
    if (tmplCol > -1 && prodCol > -1 && lastRow > 1) {
      const allRows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      const matchIdx = allRows.findIndex(r =>
        String(r[tmplCol]) === String(data['TemplateID'] || '') &&
        String(r[prodCol])  === String(data['ProductID']  || '')
      );
      if (matchIdx > -1) {
        // Update in place, preserving the original LineID
        const existingLineId = allRows[matchIdx][0];
        const rowValues = headers.map(h => {
          if (h === headers[0]) return existingLineId; // preserve original LineID
          return (data[h] !== undefined && data[h] !== null) ? data[h] : '';
        });
        sheet.getRange(matchIdx + 2, 1, 1, rowValues.length).setValues([rowValues]);
        return jsonOut({ success: true, action: 'updated' });
      }
      // Not found — insert with a clean ID (no _N suffix)
      const allIds = allRows.map(r => String(r[0]));
      const maxNum = allIds.reduce((max, id) => {
        const m = id.match(/^MTP-(\d+)$/);
        return m ? Math.max(max, parseInt(m[1], 10)) : max;
      }, 0);
      data[headers[0]] = 'MTP-' + String(maxNum + 1).padStart(3, '0');
      const rowValues = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');
      sheet.appendRow(rowValues);
      return jsonOut({ success: true, action: 'inserted' });
    }
  }

  const idValue = String(data[headers[0]] || '');
  let rowIdx = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow-1, 1).getValues().flat().map(String);
    rowIdx = ids.indexOf(idValue);
  }

  if (data._delete) {
    if (rowIdx > -1) sheet.deleteRow(rowIdx + 2);
    return jsonOut({ success: true, action: 'deleted' });
  }

  const rowValues = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');

  if (rowIdx > -1) {
    sheet.getRange(rowIdx+2, 1, 1, rowValues.length).setValues([rowValues]);
    return jsonOut({ success: true, action: 'updated' });
  } else {
    sheet.appendRow(rowValues);
    return jsonOut({ success: true, action: 'inserted' });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
