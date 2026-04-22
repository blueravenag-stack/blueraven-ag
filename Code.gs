// BLUE RAVEN AG — Google Apps Script v1.4
// Paste into script.google.com — deploy as Web App
//   Execute as: Me
//   Who has access: Anyone
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

// ── READ ─────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const result = {};
    Object.entries(TABS).forEach(([key, tabName]) => {
      const sheet = ss.getSheetByName(tabName);
      result[key] = sheet ? sheet.getDataRange().getValues() : [];
    });
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── WRITE ────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    // Accept both application/json and text/plain (no-cors sends text/plain)
    const raw = e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    const { action, table, data } = payload;

    if (action === 'write') {
      const ss      = SpreadsheetApp.openById(SHEET_ID);
      const tabName = TABS[table];
      const sheet   = ss.getSheetByName(tabName);
      if (!sheet) throw new Error('Sheet not found: ' + tabName);

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const idValue = data[headers[0]];
      const allIds  = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
        : [];
      const rowIdx  = allIds.indexOf(String(idValue));

      if (data._delete) {
        if (rowIdx > -1) sheet.deleteRow(rowIdx + 2);
      } else {
        const rowValues = headers.map(h => data[h] !== undefined ? data[h] : '');
        if (rowIdx > -1) {
          sheet.getRange(rowIdx + 2, 1, 1, rowValues.length).setValues([rowValues]);
        } else {
          sheet.appendRow(rowValues);
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
