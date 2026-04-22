// BLUE RAVEN AG — Google Apps Script v1.5
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
};

// ── ALL REQUESTS GO THROUGH doGet ────────────────────────────────────────────
// Writes use ?action=write&table=X&data=JSON
// Reads  use ?action=read
// This avoids all CORS/redirect issues with cross-origin POST requests
function doGet(e) {
  try {
    const action = e.parameter.action || 'read';

    if (action === 'read') {
      return handleRead();
    } else if (action === 'write') {
      return handleWrite(e.parameter);
    } else {
      return jsonOut({ error: 'Unknown action: ' + action });
    }
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
  const table  = params.table;
  const raw    = params.data;
  const data   = JSON.parse(decodeURIComponent(raw));

  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const tabName = TABS[table];
  const sheet   = ss.getSheetByName(tabName);
  if (!sheet) return jsonOut({ error: 'Sheet not found: ' + tabName });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol   = headers[0];
  const idValue = String(data[idCol] || '');

  // Find existing row
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
    return jsonOut({ success: true, action: 'inserted', row: lastRow + 1 });
  }
}

// Keep doPost as fallback (not used by app but harmless)
function doPost(e) {
  return jsonOut({ error: 'Use GET requests. See app.js writeRow().' });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
