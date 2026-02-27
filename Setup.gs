// ─────────────────────────────────────────────────────────────────────────────
// Setup.gs — Project initialization and templates
// ─────────────────────────────────────────────────────────────────────────────

var SYSTEM_SHEET     = '_PS_META';
var ACTIVITY_SHEET   = '_PS_ACTIVITY';
var AUTOMATIONS_SHEET = '_PS_AUTOMATIONS';
var BASELINES_SHEET  = '_PS_BASELINES';

var COL_META_ROW    = 1;  // Row 1: column type definitions (hidden)
var DATA_START_ROW  = 2;  // Rows 2+ are data

// Column type constants
var COL_TYPES = {
  TEXT:     'TEXT',
  NUMBER:   'NUMBER',
  DATE:     'DATE',
  DROPDOWN: 'DROPDOWN',
  CHECKBOX: 'CHECKBOX',
  CONTACT:  'CONTACT',
  FORMULA:  'FORMULA',
  DURATION: 'DURATION',
  PERCENT:  'PERCENT',
  AUTONUMBER: 'AUTONUMBER',
  PREDECESSOR: 'PREDECESSOR',
  ATTACHMENT: 'ATTACHMENT'
};

// Default status colours for Kanban
var STATUS_COLORS = {
  'Not Started':   '#f1f3f4',
  'In Progress':   '#4fc3f7',
  'Blocked':       '#ef5350',
  'In Review':     '#ffa726',
  'Done':          '#66bb6a',
  'Cancelled':     '#9e9e9e'
};

// ── Project templates ─────────────────────────────────────────────────────────
function setupBlankProject() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = Browser.inputBox('📋 New Project', 'Enter project name:', Browser.Buttons.OK_CANCEL);
  if (name === 'cancel' || !name) return;

  var sheet = ss.insertSheet(name);
  _initProjectSheet(sheet, [
    { name: 'Task Name',    type: COL_TYPES.TEXT,     width: 280 },
    { name: 'Assigned To',  type: COL_TYPES.CONTACT,  width: 160 },
    { name: 'Start Date',   type: COL_TYPES.DATE,     width: 110 },
    { name: 'End Date',     type: COL_TYPES.DATE,     width: 110 },
    { name: 'Duration',     type: COL_TYPES.DURATION, width: 90  },
    { name: 'Status',       type: COL_TYPES.DROPDOWN, width: 130,
      options: ['Not Started','In Progress','Blocked','In Review','Done','Cancelled'] },
    { name: '% Complete',   type: COL_TYPES.PERCENT,  width: 90  },
    { name: 'Priority',     type: COL_TYPES.DROPDOWN, width: 100,
      options: ['Low','Medium','High','Critical'] },
    { name: 'Predecessor',  type: COL_TYPES.PREDECESSOR, width: 100 },
    { name: 'Notes',        type: COL_TYPES.TEXT,     width: 240 },
    { name: 'Attachments',  type: COL_TYPES.ATTACHMENT, width: 120 }
  ]);
  ensureSystemSheets(ss);
  SpreadsheetApp.getUi().alert('✅ Project "' + name + '" created!');
}

function setupGanttProject() {
  setupBlankProject();
  // Pre-populate sample rows
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var today = new Date();
  var samples = [
    ['Phase 1: Discovery', '', today, addDays(today,7), 7, 'In Progress', 60, 'High', '', ''],
    ['  1.1 Stakeholder interviews', 'You', today, addDays(today,3), 3, 'Done', 100, 'High', '', ''],
    ['  1.2 Requirements gathering', 'You', addDays(today,2), addDays(today,7), 5, 'In Progress', 30, 'High', '2', ''],
    ['Phase 2: Design', '', addDays(today,8), addDays(today,18), 10, 'Not Started', 0, 'Medium', '', ''],
    ['  2.1 Wireframes', '', addDays(today,8), addDays(today,12), 4, 'Not Started', 0, 'Medium', '3', ''],
    ['  2.2 Design review', '', addDays(today,13), addDays(today,18), 5, 'Not Started', 0, 'Medium', '5', ''],
    ['Phase 3: Development', '', addDays(today,19), addDays(today,40), 21, 'Not Started', 0, 'High', '', ''],
  ];
  var startRow = DATA_START_ROW + 1;
  for (var i = 0; i < samples.length; i++) {
    var r = sheet.getRange(startRow + i, 1, 1, samples[i].length);
    r.setValues([samples[i]]);
    if (samples[i][0].startsWith('  ')) {
      setIndentLevel(sheet, startRow + i, 1);
    }
  }
  applyRowFormatting(sheet);
}

function setupKanbanProject() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = Browser.inputBox('🃏 New Kanban Board', 'Enter board name:', Browser.Buttons.OK_CANCEL);
  if (name === 'cancel' || !name) return;
  var sheet = ss.insertSheet(name);
  _initProjectSheet(sheet, [
    { name: 'Task Name',   type: COL_TYPES.TEXT,     width: 280 },
    { name: 'Status',      type: COL_TYPES.DROPDOWN, width: 130,
      options: ['Backlog','To Do','In Progress','In Review','Done'] },
    { name: 'Assigned To', type: COL_TYPES.CONTACT,  width: 160 },
    { name: 'Priority',    type: COL_TYPES.DROPDOWN, width: 100,
      options: ['Low','Medium','High','Critical'] },
    { name: 'Due Date',    type: COL_TYPES.DATE,     width: 110 },
    { name: 'Labels',      type: COL_TYPES.TEXT,     width: 150 },
    { name: 'Notes',       type: COL_TYPES.TEXT,     width: 240 }
  ]);
  ensureSystemSheets(ss);
  SpreadsheetApp.getUi().alert('✅ Kanban board "' + name + '" created!');
}

// ── Core sheet initializer ────────────────────────────────────────────────────
function _initProjectSheet(sheet, columns) {
  var ss = sheet.getParent();

  // Row 1: column type metadata (light grey, small font, locked to editors only)
  var metaRow = [];
  var headerRow = [];
  for (var i = 0; i < columns.length; i++) {
    var col = columns[i];
    metaRow.push(JSON.stringify({ type: col.type, options: col.options || [] }));
    headerRow.push(col.name);
  }

  // Also prepend system columns (indent level, row ID, lock flag)
  // We store these in cols A, B, C as hidden columns prefixed with _ps_
  var allMetaCols = ['{"type":"_PS_INDENT"}', '{"type":"_PS_ID"}', '{"type":"_PS_LOCKED"}'].concat(metaRow);
  var allHeaderCols = ['_indent', '_id', '_locked'].concat(headerRow);

  var totalCols = allHeaderCols.length;
  sheet.getRange(1, 1, 1, totalCols).setValues([allHeaderCols]);
  sheet.getRange(1, 1, 1, totalCols)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  // Store meta in sheet's developer metadata
  sheet.addDeveloperMetadata('PS_COLUMNS', JSON.stringify(allMetaCols));
  sheet.addDeveloperMetadata('PS_PROJECT', 'true');

  // Hide system columns (A=indent, B=id, C=locked)
  sheet.hideColumns(1, 3);

  // Set column widths for visible columns (D onwards)
  for (var j = 0; j < columns.length; j++) {
    sheet.setColumnWidth(j + 4, columns[j].width || 150);
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set up dropdown validation for DROPDOWN columns
  _applyColumnValidations(sheet, columns);

  // Auto-number the ID column on new rows via a named trigger marker
  _seedAutoIds(sheet);

  // Add conditional formatting for Status column
  _applyStatusFormatting(sheet, columns);

  // Zebra stripe data rows
  _applyAlternatingRows(sheet);
}

function _applyColumnValidations(sheet, columns) {
  for (var i = 0; i < columns.length; i++) {
    var col = columns[i];
    var colIndex = i + 4; // offset by 3 system cols + 1 for 1-based
    if (col.type === COL_TYPES.DROPDOWN && col.options && col.options.length > 0) {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(col.options, true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(DATA_START_ROW, colIndex, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
    } else if (col.type === COL_TYPES.DATE) {
      var dateRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(true)
        .build();
      sheet.getRange(DATA_START_ROW, colIndex, sheet.getMaxRows() - 1, 1).setDataValidation(dateRule);
      sheet.getRange(DATA_START_ROW, colIndex, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('dd/MM/yyyy');
    } else if (col.type === COL_TYPES.PERCENT) {
      sheet.getRange(DATA_START_ROW, colIndex, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('0"%"');
    } else if (col.type === COL_TYPES.NUMBER || col.type === COL_TYPES.DURATION) {
      sheet.getRange(DATA_START_ROW, colIndex, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('0');
    } else if (col.type === COL_TYPES.CHECKBOX) {
      var checkRule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      sheet.getRange(DATA_START_ROW, colIndex, sheet.getMaxRows() - 1, 1).setDataValidation(checkRule);
    }
  }
}

function _applyStatusFormatting(sheet, columns) {
  var statusColIndex = -1;
  for (var i = 0; i < columns.length; i++) {
    if (columns[i].name === 'Status') { statusColIndex = i + 4; break; }
  }
  if (statusColIndex < 0) return;

  var statusCol = columnToLetter(statusColIndex);
  var rules = sheet.getConditionalFormatRules();
  var statusEntries = Object.keys(STATUS_COLORS);
  for (var s = 0; s < statusEntries.length; s++) {
    var status = statusEntries[s];
    var color = STATUS_COLORS[status];
    var range = sheet.getRange(DATA_START_ROW + ':' + (sheet.getMaxRows()), statusColIndex, sheet.getMaxRows(), 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(color)
      .setFontColor(status === 'Not Started' ? '#333' : '#fff')
      .setRanges([range])
      .build());
  }
  sheet.setConditionalFormatRules(rules);
}

function _applyAlternatingRows(sheet) {
  // Banding for alternating row colours
  var dataRange = sheet.getRange(DATA_START_ROW, 1, sheet.getMaxRows() - 1, sheet.getLastColumn());
  var bandings = sheet.getBandings();
  bandings.forEach(function(b) { b.remove(); });
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function _seedAutoIds(sheet) {
  // Populate _id column (col B) with sequential IDs for existing rows
  var lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW);
  for (var r = DATA_START_ROW; r <= lastRow; r++) {
    if (!sheet.getRange(r, 2).getValue()) {
      sheet.getRange(r, 2).setValue('T-' + (r - 1));
    }
  }
}

function _applyAlternatingRows(sheet) {
  var dataRange = sheet.getRange(DATA_START_ROW, 1, Math.max(sheet.getMaxRows() - 1, 1), Math.max(sheet.getLastColumn(), 1));
  var bandings = sheet.getBandings();
  bandings.forEach(function(b) { b.remove(); });
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

// ── System sheet bootstrap ────────────────────────────────────────────────────
function ensureSystemSheets(ss) {
  [SYSTEM_SHEET, ACTIVITY_SHEET, AUTOMATIONS_SHEET, BASELINES_SHEET].forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      var s = ss.insertSheet(name);
      s.hideSheet();
      if (name === ACTIVITY_SHEET) {
        s.getRange(1,1,1,5).setValues([['Timestamp','User','Action','Sheet','Detail']]);
      }
      if (name === AUTOMATIONS_SHEET) {
        s.getRange(1,1,1,6).setValues([['ID','Trigger','Condition','Action','Target','Enabled']]);
      }
    }
  });
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function getProjectMeta(sheet) {
  try {
    var meta = sheet.getDeveloperMetadata();
    for (var i = 0; i < meta.length; i++) {
      if (meta[i].getKey() === 'PS_PROJECT') {
        return { exists: true, name: sheet.getName() };
      }
    }
  } catch(e) {}
  return { exists: false };
}

function getProjectStats(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var statusIdx = headers.indexOf('Status');
  if (statusIdx < 0) return { total: 0, done: 0, inProgress: 0, blocked: 0, pct: 0 };

  var total = 0, done = 0, inProgress = 0, blocked = 0;
  for (var i = DATA_START_ROW - 1; i < data.length; i++) {
    var row = data[i];
    if (!row[statusIdx]) continue;
    total++;
    var s = String(row[statusIdx]).toLowerCase();
    if (s === 'done') done++;
    else if (s === 'in progress') inProgress++;
    else if (s === 'blocked') blocked++;
  }
  return { total: total, done: done, inProgress: inProgress, blocked: blocked,
           pct: total > 0 ? Math.round(done / total * 100) : 0 };
}

function addDays(date, days) {
  var d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function columnToLetter(col) {
  var letter = '';
  while (col > 0) {
    var temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function applyRowFormatting(sheet) {
  _applyAlternatingRows(sheet);
}

function exportPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var blob = ss.getBlob().getAs('application/pdf');
  blob.setName(sheet.getName() + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '.pdf');
  DriveApp.createFile(blob);
  SpreadsheetApp.getUi().alert('✅ PDF saved to your Google Drive!');
}

function saveBaseline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var baselineSheet = ss.getSheetByName(BASELINES_SHEET) || ss.insertSheet(BASELINES_SHEET);
  baselineSheet.showSheet();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  baselineSheet.appendRow(['=== BASELINE: ' + sheet.getName() + ' @ ' + timestamp + ' ===']);
  data.forEach(function(row) { baselineSheet.appendRow(row); });
  baselineSheet.appendRow([]);
  baselineSheet.hideSheet();
  SpreadsheetApp.getUi().alert('✅ Baseline snapshot saved!');
}
