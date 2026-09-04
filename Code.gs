/**
 * HCL GUVI - Learner Feedback Form Backend
 * -----------------------------------------
 * Paste this into: Extensions > Apps Script (inside your Google Sheet)
 * https://docs.google.com/spreadsheets/d/1D7MdsPp20xaGJfBRQnvxFr-roaMJbTlM22DV9l-mBzY/edit
 *
 * Then Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 * Copy the resulting /exec URL into ENDPOINT in the HTML form file.
 */

var SHEET_NAME = 'Form Responses';

var HEADERS = [
  'No',
  'Name',
  'Email ID',
  'Sample Task / Project',
  'Needed Technologies',
  'Issue',
  'Description',
  'Notes',
  'Suggestions',
  'HyreX organisation ID',
  'Dev Status',
  'Dev Deadline',
  'Submitted At'
];

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#6729FF');
      headerRange.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, HEADERS.length);
    }

    var data = e.parameter;
    var nextNo = sheet.getLastRow(); // header = row 1, so lastRow before insert = next serial number

    sheet.appendRow([
      nextNo,
      data.name || '',
      data.email || '',
      data.project || '',
      data.tech || '',
      data.issue || '',
      data.description || '',
      data.notes || '',
      data.suggestions || '',
      data.hyrexId || '',
      'Pending',
      '',
      new Date()
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', row: nextNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: lets you sanity-check the deployment by visiting the /exec URL in a browser
function doGet(e) {
  return ContentService.createTextOutput('HCL GUVI form endpoint is live.');
}
