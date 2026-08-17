/**
 * Dollar Time Machine — response collector (two-tab version)
 * Paste this into Extensions → Apps Script of your Google Sheet, then
 * Deploy → Manage deployments → edit ✏ → Version: New version → Deploy.
 *
 * It writes to TWO tabs (both created automatically):
 *   • "Grid"      — one row per student, one column per turn (easy to grade).
 *                   Cell shows the student's call + ✓ / ✗ / ≈ (flat month).
 *   • "Responses" — one row per answer, keeping the full typed rationale
 *                   and timestamps.
 */

var LONG_SHEET = 'Responses';
var GRID_SHEET = 'Grid';
var LONG_HEADERS = ['Server time','Client time','Name','Session',
                    'Turn','Decision month','Choice','Actual move','Correct?','Rationale'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);                 // serialize writes so rows/cells don't clash
    var d = {};
    try { d = JSON.parse(e.postData.contents); } catch (err) { d = {}; }
    writeLong_(d);
    writeGrid_(d);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Visit the Web-app URL in a browser to confirm it's deployed and public.
function doGet() {
  return json_({ ok: true, message: 'Dollar Time Machine collector is live. POST responses here.' });
}

/* ---- long log: one row per answer (keeps rationale) ---- */
function writeLong_(d) {
  var sh = getSheet_(LONG_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(LONG_HEADERS);
    sh.getRange(1, 1, 1, LONG_HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([
    new Date(), d.clientTime || '', d.name || '', d.session || '',
    d.turn || '', d.date || '', d.choice || '', d.actual || '',
    d.correct || '', d.rationale || ''
  ]);
}

/* ---- grid: one row per student, one column per turn ---- */
function writeGrid_(d) {
  var n = turnNumber_(d);
  if (!n) return;                          // need a turn number to pick a column
  var name = String(d.name || 'Unknown').trim();
  var col = n + 1;                         // column 1 = Name, turn k => column k+1

  var sh = getSheet_(GRID_SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1).setValue('Name').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
  }
  // label the column the first time we see this turn
  var head = sh.getRange(1, col);
  if (!head.getValue()) {
    head.setValue('T' + n + (d.date ? (' · ' + d.date) : '')).setFontWeight('bold');
  }
  // find (or create) this student's row
  var row = -1, last = sh.getLastRow();
  if (last >= 2) {
    var names = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim().toLowerCase() === name.toLowerCase()) { row = i + 2; break; }
    }
  }
  if (row === -1) { row = Math.max(last, 1) + 1; sh.getRange(row, 1).setValue(name); }

  // cell = just the student's call (Buy / Sell / Flat)
  sh.getRange(row, col).setValue(d.choice || '');
}

/* ---- helpers ---- */
function turnNumber_(d) {
  var n = Number(d.n);
  if (n) return n;
  if (typeof d.turn === 'number') return d.turn;
  var m = String(d.turn || '').match(/(\d+)/);   // pull a number out of e.g. "Turn 9" or "...(#9)"
  return m ? Number(m[1]) : 0;
}

function getSheet_(nm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nm);
  if (!sh) sh = ss.insertSheet(nm);
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
