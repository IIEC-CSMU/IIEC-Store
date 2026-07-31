/**
 * Store.IIEC.in — early-access form backend.
 *
 * POST from index.html  ->  formatted Google Sheet  ->  .xlsx in Drive.
 *
 * FIRST RUN
 *  1. script.google.com -> New project -> paste this file -> Save.
 *  2. Run `setup` from the editor and approve the authorisation prompt
 *     (Sheets + Drive + external requests). This builds the header row, banding,
 *     filter, column formats, dropdowns, named range and the Summary tab, then
 *     writes the first .xlsx.
 *  3. Deploy -> New deployment -> type: Web app
 *       Execute as:      Me
 *       Who has access:  Anyone            <- required, the page is public
 *     Copy the /exec URL and paste it into FORM_ENDPOINT in index.html.
 *  4. Every later code edit needs Deploy -> Manage deployments -> pencil ->
 *     Version: NEW VERSION -> Deploy. The /exec URL serves the last PUBLISHED
 *     version, so saving alone changes nothing.
 *
 * HELPERS you can run from the editor at any time
 *    setup()           rebuild formatting + Summary, then export
 *    reformat()        re-apply table formatting over existing rows
 *    refreshXlsxNow()  regenerate the .xlsx without adding a row
 *    testSubmit()      insert one fake row end-to-end
 *    deleteTestRows()  delete rows whose email ends in @example.edu
 *    showDiagnostics() print where the data actually is
 */

/* ------------------------------ CONFIG ------------------------------ */
var CONFIG = {
  SHEET_ID:       '1gQqBtAOwNxxZ6OWI6isYSbawfkIXtvoxRuLjhwaq_J4',
  SHEET_NAME:     'Early Access',
  SUMMARY_NAME:   'Summary',
  XLSX_FOLDER_ID: '1bz6FVv2Dq-O7CA44EQYQSdtgM2ia-HmC',
  XLSX_NAME:      'Store.IIEC.in Early Access.xlsx',
  /* Optional. Paste the ID of an existing .xlsx in your Drive to have the script
     overwrite THAT file every time (from its URL: /file/d/<ID>/view).
     Leave empty and the script manages its own file named XLSX_NAME. */
  XLSX_FILE_ID:   '',
  NAMED_RANGE:    'EarlyAccessData',
  SHARED_SECRET:  '',      // '' = no token check; else must match FORM_TOKEN in index.html
  NOTIFY_EMAIL:   ''       // '' = no email alerts; else 'you@gmail.com'
};

/* Table layout. Keep HEADERS, COL and WIDTHS in step if you add a field. */
var HEADERS = ['No.', 'Timestamp', 'Name', 'Email', 'Phone', 'Branch',
               'Branch (label)', 'Year of study', 'Page', 'User agent'];
var COL = { NO:1, TS:2, NAME:3, EMAIL:4, PHONE:5, BRANCH:6, LABEL:7, YEAR:8, PAGE:9, UA:10 };
var WIDTHS = [50, 155, 190, 235, 130, 92, 205, 108, 210, 250];

/* Must mirror the <select> options in the form. */
var BRANCHES = ['CSE','IT','AIML','DS','ECE','EEE','MECH','CIVIL','CHEM','BIOTECH','MBA','OTHER'];
var YEARS    = ['1','2','3','4','other'];

/* Warm palette, matching the site. */
var THEME = { header:'#B66A3C', headerText:'#FFFFFF', line:'#D9D2C7', band:'#F6F2EA' };

/* ------------------------------ ROUTES ------------------------------ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.diag) return json(diagnostics());
  return json({ ok:true, service:'Store.IIEC.in early access', time:new Date().toISOString() });
}

/**
 * Open <your /exec URL>?diag=1 to see where the data actually is.
 * Deliberately returns counts and file metadata only — no personal data, because
 * this endpoint is public.
 */
function diagnostics() {
  var out = { ok:true, time:new Date().toISOString() };

  try {
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    out.spreadsheet = { name:ss.getName(), id:CONFIG.SHEET_ID, tabs:ss.getSheets().map(function (s) { return s.getName(); }) };
    out.sheet = sheet
      ? { name:CONFIG.SHEET_NAME, dataRows:Math.max(sheet.getLastRow() - 1, 0), columns:sheet.getLastColumn() }
      : { error:'tab "' + CONFIG.SHEET_NAME + '" not found' };
  } catch (err) {
    out.spreadsheet = { error:String(err && err.message || err) };
  }

  try {
    var folder = DriveApp.getFolderById(CONFIG.XLSX_FOLDER_ID);
    var files = [], it = folder.getFiles();
    while (it.hasNext() && files.length < 25) {
      var f = it.next();
      files.push({
        name:f.getName(), id:f.getId(), mimeType:f.getMimeType(),
        bytes:f.getSize(), updated:f.getLastUpdated().toISOString()
      });
    }
    out.folder = { name:folder.getName(), id:CONFIG.XLSX_FOLDER_ID, files:files };
  } catch (err) {
    out.folder = { error:String(err && err.message || err) };
  }

  out.trackedXlsxId = CONFIG.XLSX_FILE_ID ||
    PropertiesService.getScriptProperties().getProperty('xlsxFileId') || null;

  return out;
}

function doPost(e) {
  try {
    var data = parseBody(e);

    if (CONFIG.SHARED_SECRET && data.token !== CONFIG.SHARED_SECRET) {
      return json({ ok:false, error:'Unauthorised' });
    }

    var name   = trim(data.fullName);
    var email  = trim(data.email);
    var phone  = trim(data.phone);
    var digits = phone.replace(/\D/g, '');

    if (!name) return json({ ok:false, error:'Name is required' });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json({ ok:false, error:'Invalid email' });
    if (digits.length < 10 || digits.length > 13) return json({ ok:false, error:'Invalid phone number' });
    if (!trim(data.branch)) return json({ ok:false, error:'Branch is required' });

    var rowNumber;
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);                     // serialise concurrent submissions
    try {
      var sheet = getSheet();

      if (findRow(sheet, email, digits) > 0) {
        return json({ ok:true, duplicate:true, message:'Already on the list' });
      }

      rowNumber = sheet.getLastRow() + 1;
      sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([[
        rowNumber - 1,                                        // No.
        data.submittedAt ? new Date(data.submittedAt) : new Date(),
        titleCase(name),
        email.toLowerCase(),
        phone,
        trim(data.branch).toUpperCase(),
        trim(data.branchLabel),
        trim(data.studyYear),
        trim(data.page),
        trim(data.userAgent)
      ]]);

      styleTable(sheet);                      // keeps banding, filter and borders in step
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    // Neither of these may break the submission.
    var xlsxOk = false, xlsxId = null;
    try {
      var x = exportXlsx();
      xlsxOk = true;
      xlsxId = x.id;
      console.log('xlsx updated: ' + x.id + ' (' + x.bytes + ' bytes, created=' + x.created + ')');
    } catch (err) { console.warn('xlsx export failed: ' + err); }

    if (CONFIG.NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, 'New Store.IIEC.in early-access signup',
          [name, email, phone, trim(data.branchLabel), trim(data.studyYear)].join('\n'));
      } catch (err) { console.warn('notify failed: ' + err); }
    }

    return json({ ok:true, row:rowNumber, xlsx:xlsxOk, xlsxId:xlsxId });
  } catch (err) {
    console.error(err);
    return json({ ok:false, error:String(err && err.message || err) });
  }
}

/* ------------------------------ SHEET / TABLE ------------------------------ */

/**
 * Deletes the empty default "Sheet1" and puts the data tab first, so the
 * exported .xlsx opens on the interest list instead of a blank sheet.
 */
function tidyWorkbook() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheets = ss.getSheets();

  // drop leftover empty default tabs
  sheets.forEach(function (s) {
    var name = s.getName();
    var isDefault = /^Sheet\s?\d+$/i.test(name);
    var isOurs = (name === CONFIG.SHEET_NAME || name === CONFIG.SUMMARY_NAME);
    if (isDefault && !isOurs && s.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(s);
    }
  });

  // order: data first, summary second
  var data = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (data) { ss.setActiveSheet(data); ss.moveActiveSheet(1); }
  var summary = ss.getSheetByName(CONFIG.SUMMARY_NAME);
  if (summary) { ss.setActiveSheet(summary); ss.moveActiveSheet(2); }
  if (data) ss.setActiveSheet(data);          // leave the data tab selected
}

function getSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME, 0);
    styleTable(sheet);
  } else if (sheet.getLastRow() === 0) {
    styleTable(sheet);
  }
  ensureWidth(sheet);          // a tab built for fewer fields would break the insert
  return sheet;
}

/**
 * Makes the tab exactly HEADERS.length columns wide.
 * Without the insert branch, adding a field to HEADERS leaves the sheet too
 * narrow and every setValues() fails with "data has N but the range has M".
 */
function ensureWidth(sheet) {
  var cols = HEADERS.length;
  var have = sheet.getMaxColumns();
  if (have < cols) sheet.insertColumnsAfter(have, cols - have);
  else if (have > cols && sheet.getLastColumn() <= cols) sheet.deleteColumns(cols + 1, have - cols);
}

/**
 * Applies the whole table treatment: header row, frozen pane, alternating row
 * bands, filter, borders, column widths, number formats and dropdowns.
 * Safe to run repeatedly — it clears what it owns before re-applying.
 */
function styleTable(sheet) {
  var cols = HEADERS.length;
  var last = Math.max(sheet.getLastRow(), 1);
  var rows = last - 1;                                  // data rows, excluding header

  // match the sheet width to HEADERS: grow if short, trim the empty tail if long
  ensureWidth(sheet);

  // header
  sheet.getRange(1, 1, 1, cols)
    .setValues([HEADERS])
    .setBackground(THEME.header)
    .setFontColor(THEME.headerText)
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(1, 30);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  WIDTHS.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  // banding — the closest equivalent to an Excel table style that survives export
  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(1, 1, Math.max(last, 2), cols)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(THEME.header)
    .setFirstRowColor('#FFFFFF')
    .setSecondRowColor(THEME.band);

  // filter over the whole table
  var filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(1, 1, Math.max(last, 2), cols).createFilter();

  // formats
  sheet.getRange(2, COL.NO, Math.max(rows, 1), 1).setNumberFormat('0').setHorizontalAlignment('center');
  sheet.getRange(2, COL.TS, Math.max(rows, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(2, COL.PHONE, Math.max(rows, 1), 1).setNumberFormat('@');   // keeps +91 / leading zeros
  sheet.getRange(2, COL.BRANCH, Math.max(rows, 1), 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL.YEAR, Math.max(rows, 1), 1).setHorizontalAlignment('center');

  if (rows > 0) {
    var body = sheet.getRange(2, 1, rows, cols);
    body.setFontSize(10).setVerticalAlignment('middle').setWrap(false);
    body.setBorder(true, true, true, true, true, true, THEME.line, SpreadsheetApp.BorderStyle.SOLID);

    // dropdowns so manual edits stay consistent with the form
    sheet.getRange(2, COL.BRANCH, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(BRANCHES, true).setAllowInvalid(true).build());
    sheet.getRange(2, COL.YEAR, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(YEARS, true).setAllowInvalid(true).build());

    // renumber the No. column so it always reads 1..n after deletions
    var seq = [];
    for (var i = 0; i < rows; i++) seq.push([i + 1]);
    sheet.getRange(2, COL.NO, rows, 1).setValues(seq);
  }

  // named range: use =EarlyAccessData in formulas, exports to Excel as a defined name
  var ss = sheet.getParent();
  ss.getNamedRanges().forEach(function (nr) {
    if (nr.getName() === CONFIG.NAMED_RANGE) nr.remove();
  });
  ss.setNamedRange(CONFIG.NAMED_RANGE, sheet.getRange(1, 1, Math.max(last, 2), cols));
}

/** Returns the row index of a matching email or phone, or 0. */
function findRow(sheet, email, digits) {
  var rows = sheet.getLastRow() - 1;
  if (rows < 1) return 0;
  var values = sheet.getRange(2, COL.EMAIL, rows, 2).getValues();   // email + phone
  var mail = String(email).trim().toLowerCase();
  for (var i = 0; i < values.length; i++) {
    var rowMail  = String(values[i][0]).trim().toLowerCase();
    var rowPhone = String(values[i][1]).replace(/\D/g, '');
    if (rowMail === mail || (digits && rowPhone === digits)) return i + 2;
  }
  return 0;
}

/* ------------------------------ SUMMARY TAB ------------------------------ */
/**
 * Live counts written as formulas, so they stay correct in the Sheet and in the
 * exported .xlsx without the script having to recalculate anything.
 */
function buildSummary() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SUMMARY_NAME) || ss.insertSheet(CONFIG.SUMMARY_NAME);
  var src = "'" + CONFIG.SHEET_NAME + "'!";

  sheet.clear();
  sheet.getBandings().forEach(function (b) { b.remove(); });

  sheet.getRange('A1').setValue('Store.IIEC.in — Early Access Summary')
    .setFontSize(14).setFontWeight('bold').setFontColor(THEME.header);
  sheet.getRange('A2').setFormula('="Updated "&TEXT(NOW(),"yyyy-mm-dd hh:mm")')
    .setFontColor('#66635F').setFontSize(9);

  sheet.getRange('A4').setValue('Total signups').setFontWeight('bold');
  sheet.getRange('B4').setFormula('=COUNTA(' + src + '$D$2:$D)');
  sheet.getRange('A5').setValue('Latest submission').setFontWeight('bold');
  sheet.getRange('B5').setFormula('=IFERROR(TEXT(MAX(' + src + '$B$2:$B),"yyyy-mm-dd hh:mm"),"—")');
  sheet.getRange('A6').setValue('Signups today').setFontWeight('bold');
  sheet.getRange('B6').setFormula('=COUNTIFS(' + src + '$B$2:$B,">="&TODAY(),' + src + '$B$2:$B,"<"&TODAY()+1)');

  // by branch
  sheet.getRange('A8:B8').setValues([['Branch', 'Count']])
    .setFontWeight('bold').setBackground(THEME.header).setFontColor(THEME.headerText);
  BRANCHES.forEach(function (code, i) {
    var r = 9 + i;
    sheet.getRange(r, 1).setValue(code);
    sheet.getRange(r, 2).setFormula('=COUNTIF(' + src + '$F$2:$F,A' + r + ')');
  });
  var branchEnd = 8 + BRANCHES.length;
  sheet.getRange(branchEnd + 1, 1).setValue('Total').setFontWeight('bold');
  sheet.getRange(branchEnd + 1, 2).setFormula('=SUM(B9:B' + branchEnd + ')').setFontWeight('bold');

  // by year
  sheet.getRange('D8:E8').setValues([['Year of study', 'Count']])
    .setFontWeight('bold').setBackground(THEME.header).setFontColor(THEME.headerText);
  var yearLabels = ['1st year', '2nd year', '3rd year', '4th year', 'Other'];
  YEARS.forEach(function (code, i) {
    var r = 9 + i;
    sheet.getRange(r, 4).setValue(yearLabels[i]);
    sheet.getRange(r, 5).setFormula('=COUNTIF(' + src + '$H$2:$H,"' + code + '")');
  });
  sheet.getRange(9 + YEARS.length, 4).setValue('Not answered').setFontWeight('bold');
  sheet.getRange(9 + YEARS.length, 5)
    .setFormula('=COUNTA(' + src + '$D$2:$D)-SUM(E9:E' + (8 + YEARS.length) + ')').setFontWeight('bold');

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 30);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 80);
  sheet.getRange(8, 1, BRANCHES.length + 2, 2)
    .setBorder(true, true, true, true, true, true, THEME.line, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(8, 4, YEARS.length + 2, 2)
    .setBorder(true, true, true, true, true, true, THEME.line, SpreadsheetApp.BorderStyle.SOLID);
}

/* ------------------------------ XLSX ------------------------------ */
/**
 * Exports the whole workbook as .xlsx into the Drive folder, updating the file
 * in place so its ID and share link never change.
 */
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function exportXlsx() {
  var token = ScriptApp.getOAuthToken();
  var props = PropertiesService.getScriptProperties();

  // 1. snapshot the whole workbook as a real .xlsx
  var res = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error('Export failed ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  }
  var blob = res.getBlob().setName(CONFIG.XLSX_NAME);
  var bytes = blob.getBytes();
  if (bytes.length < 1000) throw new Error('Export returned only ' + bytes.length + ' bytes');

  // 2. work out which Drive file to overwrite
  var target = CONFIG.XLSX_FILE_ID || props.getProperty('xlsxFileId') || '';
  if (target) {
    try {
      var f = DriveApp.getFileById(target);
      if (f.isTrashed() || f.getMimeType() !== XLSX_MIME) target = '';
    } catch (err) { target = ''; }          // gone, renamed away or wrong type
  }
  if (!target) {
    var folder = DriveApp.getFolderById(CONFIG.XLSX_FOLDER_ID);
    var hits = folder.getFilesByName(CONFIG.XLSX_NAME);
    while (hits.hasNext()) {
      var hit = hits.next();
      if (hit.getMimeType() === XLSX_MIME) { target = hit.getId(); break; }
    }
  }

  // 3. overwrite in place, or create it the first time
  if (target) {
    var up = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files/' + target +
        '?uploadType=media&supportsAllDrives=true&fields=id,name,size,modifiedTime',
      {
        method: 'patch',
        contentType: XLSX_MIME,
        payload: bytes,
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    if (up.getResponseCode() >= 300) {
      throw new Error('Drive update ' + up.getResponseCode() + ': ' + up.getContentText().slice(0, 300));
    }
    props.setProperty('xlsxFileId', target);
    return { id:target, bytes:bytes.length, created:false, info:up.getContentText() };
  }

  var created = DriveApp.getFolderById(CONFIG.XLSX_FOLDER_ID).createFile(blob);
  props.setProperty('xlsxFileId', created.getId());
  return { id:created.getId(), bytes:bytes.length, created:true, url:created.getUrl() };
}

/* ------------------------------ HELPERS ------------------------------ */

function parseBody(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { /* fall through */ }
  }
  return (e && e.parameter) || {};
}

function trim(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function titleCase(s) {
  return s.replace(/\s+/g, ' ').replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------ MAINTENANCE ------------------------------ */

/** Run once after pasting this file: builds everything and grants permissions. */
function setup() {
  var sheet = getSheet();
  styleTable(sheet);
  buildSummary();
  tidyWorkbook();
  console.log('sheet ready, data rows: ' + Math.max(sheet.getLastRow() - 1, 0));
  console.log('xlsx: ' + JSON.stringify(exportXlsx()));
}

/** Re-apply table formatting over whatever rows exist now. */
function reformat() {
  styleTable(getSheet());
  buildSummary();
  tidyWorkbook();
  console.log('xlsx: ' + JSON.stringify(exportXlsx()));
}

function refreshXlsxNow() {
  console.log('xlsx: ' + JSON.stringify(exportXlsx()));
}

/** Prints exactly what ?diag=1 would return, straight into the editor log. */
function showDiagnostics() {
  console.log(JSON.stringify(diagnostics(), null, 2));
}

function testSubmit() {
  var out = doPost({ postData: { contents: JSON.stringify({
    token: CONFIG.SHARED_SECRET,
    fullName: 'test student',
    email: 'test.student@example.edu',
    phone: '9876543210',
    branch: 'CSE',
    branchLabel: 'Computer Science (CSE)',
    studyYear: '2',
    submittedAt: new Date().toISOString(),
    page: 'manual test'
  }) } });
  console.log(out.getContent());
}

/** Removes rows whose email ends in @example.edu (the test entries). */
function deleteTestRows() {
  var sheet = getSheet();
  var rows = sheet.getLastRow() - 1;
  if (rows < 1) return console.log('nothing to delete');
  var emails = sheet.getRange(2, COL.EMAIL, rows, 1).getValues();
  var removed = 0;
  for (var i = emails.length - 1; i >= 0; i--) {
    if (/@example\.edu$/i.test(String(emails[i][0]).trim())) {
      sheet.deleteRow(i + 2);
      removed++;
    }
  }
  styleTable(sheet);
  buildSummary();
  exportXlsx();
  console.log('deleted ' + removed + ' test row(s)');
}
