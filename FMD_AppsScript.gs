/**
 * FMD Microplan – Google Apps Script
 * Reads columns A, B, D, E, J. Edits columns D, E, J only.
 *
 * Setup: open the Google Sheet → Extensions → Apps Script → paste this file
 * → Deploy → New deployment → Web app → Execute as: Me → Access: Anyone.
 */

// Leave '' to use the FIRST tab of the spreadsheet, or put the exact tab name.
var SHEET_NAME = '';

// Optional passcode. If set, the web page will ask for it. '' = no passcode.
// (Recommended: the sheet contains Aadhaar numbers.)
var ACCESS_CODE = '';

var FIRST_DATA_ROW = 13;

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (ACCESS_CODE && String(body.code || '') !== ACCESS_CODE) {
      return json_({ error: 'AUTH' });
    }
    var sheet = getSheet_(body.sheetName);
    if (!sheet) return json_({ error: 'Sheet tab not found' });

    if (body.action === 'list') return json_({ data: listRows_(sheet) });
    if (body.action === 'update') return json_(updateRow_(sheet, body));
    return json_({ error: 'Unknown action: ' + body.action });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

function doGet() {
  return json_({ status: 'ok', message: 'FMD API running. Use POST.' });
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var n = name || SHEET_NAME;
  return n ? ss.getSheetByName(n) : ss.getSheets()[0];
}

function s_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(v);      // avoids 8.0E+11 display
  return String(v).trim();
}

/** A real data row = owner row (number + name) or an animal row (ear tag digits). */
function listRows_(sheet) {
  var last = sheet.getLastRow();
  if (last < FIRST_DATA_ROW) return [];
  var n = last - FIRST_DATA_ROW + 1;
  var vals = sheet.getRange(FIRST_DATA_ROW, 1, n, 10).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var a = s_(r[0]), b = s_(r[1]), d = s_(r[3]), e = s_(r[4]), j = s_(r[9]);
    var isOwner = /^\d+(\.0)?$/.test(a) && b !== '' && !/^\d+(\.0)?$/.test(b);
    var isTag = /^\d{8,}$/.test(j);
    if (!isOwner && !isTag) continue;          // skips repeated page headers / totals
    out.push({
      row: FIRST_DATA_ROW + i,
      sl: isOwner ? a.replace(/\.0$/, '') : '',
      name: isOwner ? b : '',
      aadhaar: isOwner ? d : '',
      mobile: isOwner ? e : '',
      tag: j
    });
  }
  return out;
}

function updateRow_(sheet, body) {
  var row = Number(body.row);
  var item = body.item || {};
  var old = body.old || {};
  if (!(row >= FIRST_DATA_ROW && row <= sheet.getLastRow())) {
    return { error: 'Row out of range' };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Conflict check: refuse if someone changed the row since it was loaded
    var cur = sheet.getRange(row, 1, 1, 10).getValues()[0];
    if (s_(cur[3]) !== String(old.aadhaar || '') ||
        s_(cur[4]) !== String(old.mobile || '') ||
        s_(cur[9]) !== String(old.tag || '')) {
      return { error: 'This row was changed by someone else. Please refresh and try again.' };
    }
    if (body.ownerRow) {
      sheet.getRange(row, 4).setValue(String(item.aadhaar || '').trim());
      sheet.getRange(row, 5).setValue(String(item.mobile || '').trim());
    }
    sheet.getRange(row, 10).setValue(String(item.tag || '').trim());
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
