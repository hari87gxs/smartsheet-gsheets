// ============================================================================
// Automations.gs — Rule-based automation engine for ProjectSheet
// ============================================================================
// Triggers:
//   • onEdit(e) — sheet change trigger
//   • onOpen(e) — initial setup
//   • runScheduledAutomations() — time-based trigger (every hour)
//
// Rule schema (stored in _PS_AUTOMATIONS sheet):
//   [id, name, enabled, trigger_type, trigger_config(JSON), action_type, action_config(JSON), last_run]
// ============================================================================

var AUTO_SHEET  = '_PS_AUTOMATIONS';
var TRIG_ROW    = 'ROW_ADDED';
var TRIG_CHANGE = 'FIELD_CHANGED';
var TRIG_DATE   = 'DATE_REACHED';
var TRIG_STATUS = 'STATUS_CHANGED';

var ACT_EMAIL   = 'SEND_EMAIL';
var ACT_ASSIGN  = 'SET_FIELD';
var ACT_NOTIFY  = 'ADD_COMMENT';
var ACT_MOVE    = 'CHANGE_STATUS';
var ACT_LOCK    = 'LOCK_ROW';
var ACT_WEBHOOK = 'CALL_WEBHOOK';

// ── Bootstrap ─────────────────────────────────────────────────────────────────
/**
 * Ensures the _PS_AUTOMATIONS sheet exists with correct headers.
 */
function bootstrapAutomations() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUTO_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(AUTO_SHEET);
  sheet.hideSheet();

  var headers = ['ID','Name','Enabled','Trigger Type','Trigger Config',
                 'Action Type','Action Config','Last Run','Run Count','Created By'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#0f3460').setFontColor('#e8eaed').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, headers.length, 160);

  // Seed two example automations
  _seedExampleRules(sheet);
  return sheet;
}

function _seedExampleRules(sheet) {
  var examples = [
    [_newId(), 'Notify on Blocked', 'TRUE', TRIG_STATUS,
     JSON.stringify({column:'Status', value:'Blocked'}),
     ACT_EMAIL,
     JSON.stringify({to:'{{assigned_to}}', subject:'Task Blocked: {{task_name}}',
       body:'Hi,\n\nTask "{{task_name}}" has been marked as Blocked.\nPlease review: {{sheet_url}}'}),
     '', 0, Session.getActiveUser().getEmail()],

    [_newId(), 'Auto-complete when 100%', 'TRUE', TRIG_CHANGE,
     JSON.stringify({column:'% Complete', operator:'equals', value:'100'}),
     ACT_MOVE,
     JSON.stringify({column:'Status', value:'Done'}),
     '', 0, Session.getActiveUser().getEmail()]
  ];
  sheet.getRange(2, 1, examples.length, examples[0].length).setValues(examples);
}

// ── Public API ────────────────────────────────────────────────────────────────
/**
 * Called from the menu: opens the automation builder dialog.
 */
function openAutomationBuilder() {
  var html = HtmlService.createHtmlOutputFromFile('Automations')
    .setTitle('Automation Rules')
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '⚡ Automation Rules');
}

/**
 * Returns all automation rules as an array for the dialog.
 */
function getAllAutomations() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUTO_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 10).getValues();
  return data.map(function(row) {
    return {
      id:          row[0], name:       row[1], enabled:    row[2],
      triggerType: row[3], triggerConfig: _safeParse(row[4]),
      actionType:  row[5], actionConfig:  _safeParse(row[6]),
      lastRun:     row[7], runCount:   row[8], createdBy:  row[9]
    };
  }).filter(function(r){ return r.id; });
}

/**
 * Saves (or updates) a rule from the dialog.
 */
function saveAutomation(rule) {
  var sheet = bootstrapAutomations();
  var existing = _findRuleRow(sheet, rule.id);
  var row = [
    rule.id, rule.name, rule.enabled,
    rule.triggerType,  JSON.stringify(rule.triggerConfig),
    rule.actionType,   JSON.stringify(rule.actionConfig),
    existing ? sheet.getRange(existing, 8).getValue() : '',
    existing ? sheet.getRange(existing, 9).getValue() : 0,
    Session.getActiveUser().getEmail()
  ];
  if (existing) {
    sheet.getRange(existing, 1, 1, row.length).setValues([row]);
  } else {
    if (!rule.id) rule.id = _newId();
    row[0] = rule.id;
    sheet.appendRow(row);
  }
  return rule.id;
}

/**
 * Deletes a rule by ID.
 */
function deleteAutomation(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUTO_SHEET);
  if (!sheet) return;
  var row = _findRuleRow(sheet, id);
  if (row) sheet.deleteRow(row);
}

/**
 * Toggles enable/disable for a rule.
 */
function toggleAutomation(id, enabled) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUTO_SHEET);
  if (!sheet) return;
  var row = _findRuleRow(sheet, id);
  if (row) sheet.getRange(row, 3).setValue(enabled);
}

// ── Trigger handlers ──────────────────────────────────────────────────────────
/**
 * Called by onEdit installable trigger. Runs FIELD_CHANGED and STATUS_CHANGED rules.
 */
function handleOnEditAutomations(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName().startsWith('_PS_')) return;      // ignore system sheets

  var rules = getAllAutomations().filter(function(r){
    return String(r.enabled).toLowerCase() === 'true' &&
           (r.triggerType === TRIG_CHANGE || r.triggerType === TRIG_STATUS);
  });
  if (!rules.length) return;

  var headers = _getHeaders(sheet);
  var colIndex = e.range.getColumn() - 1; // 0-based
  var changedCol = headers[colIndex] || '';
  var newValue   = e.value || '';
  var rowIndex   = e.range.getRow();

  rules.forEach(function(rule) {
    var tc = rule.triggerConfig;
    if (!tc || tc.column !== changedCol) return;

    var matches = false;
    if (rule.triggerType === TRIG_STATUS) {
      matches = newValue === tc.value;
    } else if (rule.triggerType === TRIG_CHANGE) {
      matches = _evalCondition(newValue, tc.operator, tc.value);
    }
    if (matches) _executeAction(rule, sheet, rowIndex, headers);
  });
}

/**
 * Called by onEdit — also fires ROW_ADDED when first content cell in a new row is set.
 */
function handleRowAddedAutomations(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName().startsWith('_PS_')) return;

  // Heuristic: new row = row > lastRow-1 and column D (first user col) just set
  if (e.range.getColumn() !== 4 || !e.value) return;

  var rules = getAllAutomations().filter(function(r){
    return String(r.enabled).toLowerCase() === 'true' && r.triggerType === TRIG_ROW;
  });
  if (!rules.length) return;

  var headers = _getHeaders(sheet);
  rules.forEach(function(rule){
    _executeAction(rule, sheet, e.range.getRow(), headers);
  });
}

/**
 * Scheduled trigger (every hour). Evaluates DATE_REACHED rules.
 */
function runScheduledAutomations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rules = getAllAutomations().filter(function(r){
    return String(r.enabled).toLowerCase() === 'true' && r.triggerType === TRIG_DATE;
  });
  if (!rules.length) return;

  var today = new Date(); today.setHours(0,0,0,0);

  ss.getSheets().filter(function(sh){return !sh.getName().startsWith('_PS_');}).forEach(function(sheet){
    var headers = _getHeaders(sheet);
    if (!sheet.getLastRow() || sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();

    rules.forEach(function(rule) {
      var tc = rule.triggerConfig;
      var col = headers.indexOf(tc.dateColumn);
      if (col < 0) return;

      data.forEach(function(row, i) {
        var cellDate = row[col];
        if (!cellDate) return;
        var d = new Date(cellDate); d.setHours(0,0,0,0);
        var diff = Math.round((d - today)/(1000*60*60*24));
        var shouldFire = (tc.when === 'on_date' && diff === 0) ||
                         (tc.when === 'days_before' && diff === parseInt(tc.days||0)) ||
                         (tc.when === 'days_after'  && diff === -parseInt(tc.days||0));
        if (shouldFire) _executeAction(rule, sheet, i+2, headers);
      });
    });
  });
}

// ── Action executor ───────────────────────────────────────────────────────────
function _executeAction(rule, sheet, rowIndex, headers) {
  try {
    var ac  = rule.actionConfig;
    var ctx = _buildContext(sheet, rowIndex, headers);

    if (rule.actionType === ACT_EMAIL) {
      var to      = _interpolate(ac.to, ctx);
      var subject = _interpolate(ac.subject, ctx);
      var body    = _interpolate(ac.body, ctx);
      if (to && _isEmail(to)) {
        GmailApp.sendEmail(to, subject, body);
        logActivity(sheet, rowIndex, 'AUTO_EMAIL', 'Sent email: ' + subject);
      }

    } else if (rule.actionType === ACT_MOVE) {
      var colIdx = headers.indexOf(ac.column);
      if (colIdx >= 0) {
        sheet.getRange(rowIndex, colIdx+1).setValue(ac.value);
        logActivity(sheet, rowIndex, 'AUTO_STATUS', ac.column + ' → ' + ac.value);
      }

    } else if (rule.actionType === ACT_ASSIGN) {
      var colIdx2 = headers.indexOf(ac.column);
      if (colIdx2 >= 0) {
        sheet.getRange(rowIndex, colIdx2+1).setValue(_interpolate(ac.value, ctx));
        logActivity(sheet, rowIndex, 'AUTO_FIELD', 'Set ' + ac.column + ' = ' + ac.value);
      }

    } else if (rule.actionType === ACT_NOTIFY) {
      var commentCol = headers.indexOf('Notes');
      if (commentCol >= 0) {
        var existing = sheet.getRange(rowIndex, commentCol+1).getValue();
        var note = '[Auto ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm') + '] ' +
                   _interpolate(ac.message, ctx);
        sheet.getRange(rowIndex, commentCol+1).setValue(existing ? existing + '\n' + note : note);
      }

    } else if (rule.actionType === ACT_LOCK) {
      lockRow(sheet, rowIndex);

    } else if (rule.actionType === ACT_WEBHOOK) {
      var url     = _interpolate(ac.url, ctx);
      var payload = JSON.stringify(_interpolate(ac.payload || '{}', ctx));
      UrlFetchApp.fetch(url, {method:'post', contentType:'application/json', payload:payload,
        muteHttpExceptions:true});
    }

    // Update run count & last run
    var autoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUTO_SHEET);
    if (autoSheet) {
      var ruleRow = _findRuleRow(autoSheet, rule.id);
      if (ruleRow) {
        autoSheet.getRange(ruleRow, 8).setValue(new Date());
        autoSheet.getRange(ruleRow, 9).setValue((parseInt(autoSheet.getRange(ruleRow,9).getValue())||0)+1);
      }
    }
  } catch(err) {
    Logger.log('[Automation Error] Rule: ' + rule.name + ' | ' + err.message);
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function _buildContext(sheet, rowIndex, headers) {
  var rowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  var ctx = {};
  headers.forEach(function(h,i){ ctx[h] = rowData[i]; });
  ctx['task_name']   = rowData[headers.indexOf('Task Name')] || rowData[3] || '';
  ctx['assigned_to'] = rowData[headers.indexOf('Assigned To')] || '';
  ctx['status']      = rowData[headers.indexOf('Status')] || '';
  ctx['sheet_url']   = sheet.getParent().getUrl();
  ctx['sheet_name']  = sheet.getName();
  return ctx;
}

function _interpolate(template, ctx) {
  if (!template) return '';
  return String(template).replace(/\{\{(\w+)\}\}/g, function(_, key){ return ctx[key] || ''; });
}

function _evalCondition(value, operator, target) {
  var v = parseFloat(value), t = parseFloat(target);
  switch(operator) {
    case 'equals':          return String(value) === String(target);
    case 'not_equals':      return String(value) !== String(target);
    case 'contains':        return String(value).toLowerCase().includes(String(target).toLowerCase());
    case 'greater_than':    return !isNaN(v) && !isNaN(t) && v > t;
    case 'less_than':       return !isNaN(v) && !isNaN(t) && v < t;
    case 'is_empty':        return !value;
    case 'is_not_empty':    return !!value;
    default: return false;
  }
}

function _getHeaders(sheet) {
  if (sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}

function _findRuleRow(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return null;
  var ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
  for (var i=0;i<ids.length;i++) { if (ids[i][0] === id) return i+2; }
  return null;
}

function _safeParse(s) {
  try { return JSON.parse(s); } catch(e) { return {}; }
}

function _isEmail(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }
function _newId()    { return 'auto_' + Date.now(); }

// ── Time trigger installation ─────────────────────────────────────────────────
/**
 * Install or remove the hourly scheduled trigger.
 */
function installScheduledTrigger() {
  // Remove existing
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'runScheduledAutomations') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runScheduledAutomations').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('✅ Hourly automation trigger installed.');
}

function removeScheduledTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'runScheduledAutomations') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert('Scheduled automation trigger removed.');
}
