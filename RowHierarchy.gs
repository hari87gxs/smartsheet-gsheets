// ─────────────────────────────────────────────────────────────────────────────
// RowHierarchy.gs — Parent/child row indentation (like Smartsheet row hierarchy)
// ─────────────────────────────────────────────────────────────────────────────

var INDENT_COL   = 1;  // Hidden col A stores indent level (0 = root)
var INDENT_CHARS = '    '; // 4 spaces per indent level in display col

function indentRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row <= COL_META_ROW) return;
  var current = Number(sheet.getRange(row, INDENT_COL).getValue()) || 0;
  setIndentLevel(sheet, row, current + 1);
  logActivity('INDENT', 'Row ' + row + ' indented to level ' + (current + 1));
}

function outdentRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row <= COL_META_ROW) return;
  var current = Number(sheet.getRange(row, INDENT_COL).getValue()) || 0;
  if (current === 0) return;
  setIndentLevel(sheet, row, current - 1);
  logActivity('OUTDENT', 'Row ' + row + ' outdented to level ' + (current - 1));
}

function setIndentLevel(sheet, row, level) {
  sheet.getRange(row, INDENT_COL).setValue(level);
  _applyIndentFormatting(sheet, row, level);
}

function _applyIndentFormatting(sheet, row, level) {
  var taskNameCol = 4; // First visible col (after 3 hidden system cols)
  var cell = sheet.getRange(row, taskNameCol);
  var padding = '';
  for (var i = 0; i < level; i++) padding += INDENT_CHARS;
  var raw = String(cell.getValue()).replace(/^\s+/, '');
  cell.setValue(padding + raw);

  if (level === 0) {
    // Root row: bold
    cell.setFontWeight('bold').setFontSize(11).setBackground(null);
    sheet.getRange(row, taskNameCol, 1, sheet.getLastColumn() - taskNameCol + 1)
      .setBackground('#e8f0fe');
  } else if (level === 1) {
    // First level child: normal weight
    cell.setFontWeight('normal').setFontSize(10);
  } else {
    // Deeper children: italic, lighter
    cell.setFontStyle('italic').setFontSize(10).setFontColor('#555555');
  }
}

function collapseChildren() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var parentLevel = Number(sheet.getRange(row, INDENT_COL).getValue()) || 0;
  var lastRow = sheet.getLastRow();
  var toHide = [];
  for (var r = row + 1; r <= lastRow; r++) {
    var level = Number(sheet.getRange(r, INDENT_COL).getValue()) || 0;
    if (level <= parentLevel) break;
    toHide.push(r);
  }
  if (toHide.length > 0) {
    sheet.hideRows(toHide[0], toHide.length);
  }
}

function expandChildren() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var parentLevel = Number(sheet.getRange(row, INDENT_COL).getValue()) || 0;
  var lastRow = sheet.getLastRow();
  for (var r = row + 1; r <= lastRow; r++) {
    var level = Number(sheet.getRange(r, INDENT_COL).getValue()) || 0;
    if (level <= parentLevel) break;
    sheet.showRows(r);
  }
}

// ── Get all rows as a tree structure (used by Gantt, Kanban, Dashboard) ───────
function getRowTree(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = DATA_START_ROW - 1; i < data.length; i++) {
    var row = data[i];
    if (!row[3] && !row[4]) continue; // skip totally empty rows
    var obj = { _rowIndex: i + 1, _indent: Number(row[INDENT_COL - 1]) || 0 };
    for (var j = 3; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    rows.push(obj);
  }
  return rows;
}

// ── Row locking ───────────────────────────────────────────────────────────────
function lockSelectedRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var protection = range.protect().setDescription('Locked row');
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
  // Mark in system col C
  for (var r = startRow; r < startRow + numRows; r++) {
    sheet.getRange(r, 3).setValue('LOCKED');
  }
  logActivity('LOCK', 'Rows ' + startRow + '-' + (startRow + numRows - 1) + ' locked');
  SpreadsheetApp.getUi().alert('✅ Rows locked.');
}
