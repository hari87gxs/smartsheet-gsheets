// ============================================================================
// ActivityLog.gs — Change tracking, row comments, and @mention notifications
// ============================================================================

var ACTIVITY_SHEET = '_PS_ACTIVITY';
var MAX_LOG_ROWS   = 5000;          // auto-trim after this many entries

// ── Log schema ─────────────────────────────────────────────────────────────────
//  [Timestamp, Sheet, Row, User, Action, Detail, Old Value, New Value]
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Ensures the activity log sheet exists.
 */
function bootstrapActivityLog() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ACTIVITY_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(ACTIVITY_SHEET);
  sheet.hideSheet();

  var headers = ['Timestamp','Sheet','Row','User','Action','Detail','Old Value','New Value'];
  sheet.getRange(1,1,1,headers.length).setValues([headers])
    .setBackground('#0f3460').setFontColor('#e8eaed').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(6, 300);
  return sheet;
}

/**
 * Write one activity record.
 *
 * @param {Sheet}  sheet     — the project sheet where the change happened
 * @param {number} rowIndex  — 1-based row number (0 = sheet-level action)
 * @param {string} action    — e.g. 'EDIT', 'STATUS_CHANGE', 'COMMENT', 'AUTO_EMAIL'
 * @param {string} detail    — human-readable description
 * @param {*}      oldValue  — (optional) previous value
 * @param {*}      newValue  — (optional) new value
 */
function logActivity(sheet, rowIndex, action, detail, oldValue, newValue) {
  try {
    var log = bootstrapActivityLog();
    var user = Session.getActiveUser().getEmail() || 'system';
    log.appendRow([
      new Date(),
      sheet ? sheet.getName() : '',
      rowIndex || '',
      user,
      action || '',
      detail || '',
      oldValue !== undefined ? String(oldValue) : '',
      newValue !== undefined ? String(newValue) : ''
    ]);

    // Auto-trim
    if (log.getLastRow() > MAX_LOG_ROWS + 1) {
      log.deleteRows(2, log.getLastRow() - MAX_LOG_ROWS - 1);
    }
  } catch(err) {
    Logger.log('[ActivityLog Error] ' + err.message);
  }
}

/**
 * Called by the onEdit trigger to log every cell change automatically.
 */
function handleOnEditActivityLog(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName().startsWith('_PS_')) return;
  if (e.range.getRow() === 1) return;          // header edit, skip

  var headers = sheet.getLastColumn() > 0
    ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
  var colName  = headers[e.range.getColumn()-1] || ('Col ' + e.range.getColumn());
  var detail   = colName + ' changed';

  logActivity(sheet, e.range.getRow(), 'EDIT', detail, e.oldValue, e.value);
}

// ── Activity viewer dialog ────────────────────────────────────────────────────
/**
 * Opens the activity log viewer (launched from menu).
 */
function openActivityLog() {
  var html = HtmlService.createHtmlOutput(buildActivityLogHtml())
    .setTitle('Activity Log')
    .setWidth(900)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Activity Log');
}

function buildActivityLogHtml() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVITY_SHEET);
  var rows  = [];
  if (sheet && sheet.getLastRow() > 1) {
    rows = sheet.getRange(2, 1, Math.min(sheet.getLastRow()-1, 200), 8).getValues()
      .reverse();   // newest first
  }

  var rowsHtml = rows.map(function(r) {
    var ts      = r[0] instanceof Date ? r[0].toLocaleString('en-SG') : String(r[0]);
    var action  = String(r[4]);
    var color   = {EDIT:'#4fc3f7',STATUS_CHANGE:'#ffa726',COMMENT:'#66bb6a',
                   AUTO_EMAIL:'#9c27b0',AUTO_STATUS:'#7e57c2',COMMENT_MENTION:'#ef5350'}[action] || '#9aa0a6';
    return '<tr>' +
      '<td style="color:#9aa0a6;white-space:nowrap">'+esc(ts)+'</td>' +
      '<td>'+esc(String(r[1]))+'</td>' +
      '<td style="text-align:center">'+esc(String(r[2]))+'</td>' +
      '<td style="color:#9aa0a6">'+esc(String(r[3]).split('@')[0])+'</td>' +
      '<td><span style="background:'+color+'22;color:'+color+';padding:2px 8px;border-radius:10px;font-size:11px">'+esc(action)+'</span></td>' +
      '<td>'+esc(String(r[5]))+'</td>' +
      '<td style="color:#9aa0a6">'+esc(String(r[6]).slice(0,30))+'</td>' +
      '<td style="color:#e8eaed">'+esc(String(r[7]).slice(0,30))+'</td>' +
      '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>*{box-sizing:border-box;margin:0;padding:0;}' +
    'body{font-family:-apple-system,sans-serif;background:#1a1a2e;color:#e8eaed;font-size:12px;}' +
    '#bar{display:flex;align-items:center;gap:10px;padding:10px 16px;background:#16213e;border-bottom:1px solid #0f3460;}' +
    '#bar h2{font-size:14px;color:#4fc3f7;margin-right:auto;}' +
    'input{background:#0f3460;border:1px solid #1a73e8;color:#e8eaed;padding:5px 10px;border-radius:6px;width:200px;}' +
    '.tb-btn{background:#0f3460;border:none;color:#e8eaed;padding:6px 14px;border-radius:6px;cursor:pointer;}' +
    '.tb-btn:hover{background:#1a73e8;}' +
    '#wrap{overflow:auto;height:calc(100vh - 50px);padding:0;}' +
    'table{width:100%;border-collapse:collapse;}' +
    'th{position:sticky;top:0;background:#16213e;color:#9aa0a6;padding:8px 10px;text-align:left;font-size:11px;border-bottom:1px solid #0f3460;}' +
    'td{padding:7px 10px;border-bottom:1px solid #0f3460;vertical-align:top;}' +
    'tr:hover td{background:#16213e;}' +
    '</style></head><body>' +
    '<div id="bar"><h2>📋 Activity Log</h2>' +
    '<input id="q" placeholder="🔍 Filter…" oninput="filterLog()">' +
    '<button class="tb-btn" onclick="google.script.host.close()">✕ Close</button></div>' +
    '<div id="wrap"><table><thead><tr>' +
    '<th>Time</th><th>Sheet</th><th>Row</th><th>User</th><th>Action</th><th>Detail</th><th>Old Value</th><th>New Value</th>' +
    '</tr></thead><tbody id="body">' + rowsHtml + '</tbody></table></div>' +
    '<script>function filterLog(){var q=document.getElementById("q").value.toLowerCase();' +
    'document.querySelectorAll("#body tr").forEach(function(r){' +
    'r.style.display=r.textContent.toLowerCase().includes(q)?"":"none";});}</scr' + 'ipt>' +
    '</body></html>';
}

function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
