// ============================================================================
// Utils.gs — Server-side helpers used by HTML views via google.script.run
// ============================================================================

var DATA_START_ROW   = 2;
var SYSTEM_SHEET     = '_PS_META';

// ── getRowTree ────────────────────────────────────────────────────────────────
/**
 * Returns all data rows in the active sheet as an array of plain objects.
 * Each object has keys = column headers, plus:
 *   _rowIndex  (1-based row number)
 *   _indent    (from hidden col A)
 *   _id        (from hidden col B)
 *   _locked    (from hidden col C)
 *
 * @param {Sheet} [sheet] — defaults to active sheet
 */
function getRowTree(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  if (!sheet || sheet.getLastRow() < DATA_START_ROW) return [];

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var data    = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol).getValues();

  return data.map(function(row, i) {
    var obj = { _rowIndex: i + DATA_START_ROW, _indent: row[0] || 0,
                _id: row[1] || '', _locked: row[2] || false };
    for (var c = 3; c < headers.length; c++) {
      obj[headers[c]] = row[c];
    }
    return obj;
  }).filter(function(r) {
    // Filter out completely empty rows
    return Object.keys(r).some(function(k) { return !k.startsWith('_') && r[k] !== ''; });
  });
}

// ── updateTaskStatus ──────────────────────────────────────────────────────────
/**
 * Sets the Status column of a specific row (used by Kanban drag-drop).
 *
 * @param {number} rowIndex — 1-based
 * @param {string} newStatus
 */
function updateTaskStatus(rowIndex, newStatus) {
  var sheet   = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var colIdx  = headers.indexOf('Status');
  if (colIdx < 0) return false;
  var old = sheet.getRange(rowIndex, colIdx + 1).getValue();
  sheet.getRange(rowIndex, colIdx + 1).setValue(newStatus);
  logActivity(sheet, rowIndex, 'STATUS_CHANGE', 'Status changed', old, newStatus);
  return true;
}

// ── updateCellValue ───────────────────────────────────────────────────────────
/**
 * Generic cell update used by views.
 *
 * @param {number} rowIndex
 * @param {string} columnName
 * @param {*}      newValue
 */
function updateCellValue(rowIndex, columnName, newValue) {
  var sheet   = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var colIdx  = headers.indexOf(columnName);
  if (colIdx < 0) return false;
  var old = sheet.getRange(rowIndex, colIdx + 1).getValue();
  sheet.getRange(rowIndex, colIdx + 1).setValue(newValue);
  logActivity(sheet, rowIndex, 'EDIT', columnName + ' updated', old, newValue);
  return true;
}

// ── addTaskToSheet ────────────────────────────────────────────────────────────
/**
 * Adds a new task row at the end of the active sheet (used by Kanban + view).
 *
 * @param {string} taskName
 * @param {string} [status]   — defaults to 'To Do'
 * @param {Object} [extras]   — any extra column values {columnName: value}
 */
function addTaskToSheet(taskName, status, extras) {
  var sheet   = SpreadsheetApp.getActiveSheet();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

  var newRow  = new Array(lastCol).fill('');
  newRow[0] = 0;  // indent
  newRow[1] = _newId();  // id

  var taskCol   = headers.indexOf('Task Name');
  var statusCol = headers.indexOf('Status');
  if (taskCol   >= 0) newRow[taskCol]   = taskName;
  if (statusCol >= 0) newRow[statusCol] = status || 'To Do';

  if (extras) {
    Object.keys(extras).forEach(function(k) {
      var idx = headers.indexOf(k);
      if (idx >= 0) newRow[idx] = extras[k];
    });
  }

  sheet.appendRow(newRow);
  var newRowIndex = sheet.getLastRow();
  logActivity(sheet, newRowIndex, 'ROW_ADDED', 'New task: ' + taskName);
  return newRowIndex;
}

// ── lockRow ───────────────────────────────────────────────────────────────────
/**
 * Protects a row from editing.
 */
function lockRow(sheet, rowIndex) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  var protection = range.protect().setDescription('Locked row ' + rowIndex);
  protection.addEditor(Session.getActiveUser().getEmail());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
  // Mark col C
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var lockedCol = headers.indexOf('_locked');
  if (lockedCol >= 0) sheet.getRange(rowIndex, lockedCol+1).setValue(true);
  logActivity(sheet, rowIndex, 'LOCK', 'Row locked');
}

// ── getProjectMeta ────────────────────────────────────────────────────────────
/**
 * Returns project metadata for the homepage card.
 */
function getProjectMeta(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  var meta = sheet.getDeveloperMetadata();
  var isProject = meta.some(function(m){ return m.getKey() === 'PS_PROJECT'; });
  return {
    exists: isProject,
    name:   sheet.getName(),
    sheetId: sheet.getSheetId()
  };
}

// ── getProjectStats ───────────────────────────────────────────────────────────
/**
 * Returns summary stats for the homepage card.
 */
function getProjectStats(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  var rows = getRowTree(sheet);
  var done     = rows.filter(function(r){ return /done|complete/i.test(String(r['Status']||'')); }).length;
  var inProg   = rows.filter(function(r){ return /progress|review/i.test(String(r['Status']||'')); }).length;
  var blocked  = rows.filter(function(r){ return /blocked/i.test(String(r['Status']||'')); }).length;
  var pct      = rows.length ? Math.round(done/rows.length*100) : 0;
  return { total: rows.length, done: done, inProgress: inProg, blocked: blocked, pct: pct };
}

// ── Stubs for menu items not yet detailed ─────────────────────────────────────
function openAutomationList() { openAutomationBuilder(); }

function runAllAutomations() {
  runScheduledAutomations();
  SpreadsheetApp.getUi().alert('✅ Automations executed. Check Activity Log for results.');
}

function addRowComment() { openCommentPanel(); }

function openShareDialog() {
  var email = SpreadsheetApp.getUi().prompt(
    '🔗 Share Spreadsheet',
    'Enter email address to share with:',
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (email.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;
  var result = shareWithUser(email.getResponseText().trim(), 'editor');
  SpreadsheetApp.getUi().alert(result ? '✅ Shared successfully.' : '❌ Could not share. Check the email address.');
}

function exportPDF() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl().replace('/edit', '/export?format=pdf&portrait=false&size=A4');
  SpreadsheetApp.getUi().alert('📄 Export URL:\n\n' + url + '\n\nOpen this URL in your browser to download the PDF.');
}

function exportGanttPDF() {
  openGanttView();
  SpreadsheetApp.getUi().alert('Tip: In the Gantt chart, press Ctrl/Cmd+P to print/save as PDF.');
}

function saveBaseline() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var src   = SpreadsheetApp.getActiveSheet();
  var bName = '_PS_BASELINES';
  var base  = ss.getSheetByName(bName) || ss.insertSheet(bName);
  base.hideSheet();

  var ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var dest = ss.insertSheet('Baseline_' + ts.replace(/[: ]/g,'_'));
  src.copyTo(ss);  // copies to last sheet; move it under baselines conceptually
  logActivity(src, 0, 'BASELINE', 'Baseline saved: ' + ts);
  SpreadsheetApp.getUi().alert('✅ Baseline "' + ts + '" saved as a new sheet.');
}

// ── Install triggers ──────────────────────────────────────────────────────────
/**
 * Call once from the menu (or run manually) to install the onEdit + onOpen triggers.
 */
function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Remove duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (['_onEditDispatch','_onOpenDispatch'].includes(fn)) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('_onEditDispatch').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('_onOpenDispatch').forSpreadsheet(ss).onOpen().create();
  installScheduledTrigger();
  SpreadsheetApp.getUi().alert('✅ All triggers installed (onEdit, onOpen, hourly automations).');
}

function _onEditDispatch(e) {
  handleOnEditActivityLog(e);
  handleOnEditAutomations(e);
  handleRowAddedAutomations(e);
}

function _onOpenDispatch(e) {
  onOpen(e);
  bootstrapActivityLog();
  bootstrapAutomations();
}

// ── Internal helpers ──────────────────────────────────────────────────────────
function _newId() {
  return 'row_' + Date.now() + '_' + Math.floor(Math.random()*1000);
}
