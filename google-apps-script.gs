// Deployed Apps Script source — paste into script.google.com, bound to the
// Google Sheet that should receive submissions. Deploy as a Web App:
//   Execute as: Me
//   Who has access: Anyone
// Then paste the resulting /exec URL into APPS_SCRIPT_URL in index.html.
//
// Each form type gets its own tab. Tabs are auto-created on first submission
// with the header row defined below. To add a new form type, add an entry here
// and add a matching <form> with a hidden formType input to index.html.

const SHEETS = {
  hiring:  ['Timestamp', 'Name', 'Email', 'Company', 'LinkedIn', 'Message'],
  leader:  ['Timestamp', 'Name', 'Email', 'LinkedIn', 'Preferences'],
  advisor: ['Timestamp', 'Name', 'Email', 'LinkedIn'],
};

function doPost(e) {
  try {
    const p = e.parameter || {};
    const type = p.formType;
    if (!SHEETS[type]) {
      return json({ ok: false, error: 'invalid formType' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(type);
    if (!sheet) {
      sheet = ss.insertSheet(type);
      sheet.appendRow(SHEETS[type]);
      sheet.getRange(1, 1, 1, SHEETS[type].length).setFontWeight('bold');
    }

    const row = SHEETS[type].map(function(h) {
      if (h === 'Timestamp') return new Date();
      return p[h.toLowerCase()] || '';
    });
    sheet.appendRow(row);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, message: 'Firsthand Partners form endpoint is live.' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
