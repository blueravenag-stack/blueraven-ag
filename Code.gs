// BLUE RAVEN AG — Google Apps Script v1.6
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

// ── ALL REQUESTS GO THROUGH doGet ────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || 'read';
    if (action === 'read')  return handleRead();
    if (action === 'write') return handleWrite(e.parameter);
    return jsonOut({ error: 'Unknown action: ' + action });
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

  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const tabName = TABS[table];
  const sheet   = ss.getSheetByName(tabName);
  if (!sheet) return jsonOut({ error: 'Sheet not found: ' + tabName });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idValue = String(data[headers[0]] || '');
  const lastRow = sheet.getLastRow();

  let rowIdx = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
    rowIdx = ids.indexOf(idValue);
  }

  if (data._delete) {
    if (rowIdx > -1) sheet.deleteRow(rowIdx + 2);
    return jsonOut({ success: true, action: 'deleted' });
  }

  const rowValues = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');

  if (rowIdx > -1) {
    sheet.getRange(rowIdx + 2, 1, 1, rowValues.length).setValues([rowValues]);
    return jsonOut({ success: true, action: 'updated', row: rowIdx + 2 });
  } else {
    sheet.appendRow(rowValues);
    return jsonOut({ success: true, action: 'inserted' });
  }
}

function doPost(e) {
  return jsonOut({ error: 'Use GET requests. See app.js writeRow().' });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
