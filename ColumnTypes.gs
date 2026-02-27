// ─────────────────────────────────────────────────────────────────────────────
// ColumnTypes.gs — Column type management, dropdowns, typed input
// ─────────────────────────────────────────────────────────────────────────────

function addTypedColumn() {
  var html = HtmlService.createHtmlOutput(getAddColumnHTML())
    .setWidth(420).setHeight(380).setTitle('Add Column');
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Add Column with Type');
}

function setColumnType() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var col = sheet.getActiveRange().getColumn();
  var header = sheet.getRange(1, col).getValue();
  var html = HtmlService.createHtmlOutput(getSetColumnTypeHTML(header, col))
    .setWidth(420).setHeight(320).setTitle('Set Column Type');
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ Set Column Type: ' + header);
}

function manageDropdowns() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var col = sheet.getActiveRange().getColumn();
  var header = sheet.getRange(1, col).getValue();
  var html = HtmlService.createHtmlOutput(getManageDropdownHTML(header, col))
    .setWidth(420).setHeight(400).setTitle('Manage Dropdown');
  SpreadsheetApp.getUi().showModalDialog(html, '🔽 Manage Dropdown: ' + header);
}

// Called from Add Column dialog
function createColumn(name, type, options) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, lastCol).setValue(name)
    .setBackground('#1a1a2e').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  if (type === COL_TYPES.DROPDOWN && options && options.length > 0) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true).setAllowInvalid(false).build();
    sheet.getRange(DATA_START_ROW, lastCol, sheet.getMaxRows(), 1).setDataValidation(rule);
    _applyDropdownColors(sheet, lastCol, options);
  } else if (type === COL_TYPES.DATE) {
    sheet.getRange(DATA_START_ROW, lastCol, sheet.getMaxRows(), 1).setNumberFormat('dd/MM/yyyy');
  } else if (type === COL_TYPES.CHECKBOX) {
    sheet.getRange(DATA_START_ROW, lastCol, sheet.getMaxRows(), 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  } else if (type === COL_TYPES.PERCENT) {
    sheet.getRange(DATA_START_ROW, lastCol, sheet.getMaxRows(), 1).setNumberFormat('0"%"');
  }
  logActivity('COL_ADD', 'Column "' + name + '" (' + type + ') added');
  return 'ok';
}

function updateDropdownOptions(col, options) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true).setAllowInvalid(false).build();
  sheet.getRange(DATA_START_ROW, col, sheet.getMaxRows(), 1).setDataValidation(rule);
  _applyDropdownColors(sheet, col, options);
  logActivity('COL_UPDATE', 'Dropdown column ' + col + ' updated');
  return 'ok';
}

function _applyDropdownColors(sheet, col, options) {
  var rules = sheet.getConditionalFormatRules() || [];
  var palette = ['#4fc3f7','#66bb6a','#ffa726','#ef5350','#ab47bc','#26c6da','#ffee58','#78909c'];
  options.forEach(function(opt, i) {
    var range = sheet.getRange(DATA_START_ROW, col, sheet.getMaxRows() - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(opt)
      .setBackground(palette[i % palette.length])
      .setFontColor('#ffffff')
      .setRanges([range])
      .build());
  });
  sheet.setConditionalFormatRules(rules);
}

// ── HTML for dialogs ──────────────────────────────────────────────────────────
function getAddColumnHTML() {
  return '<!DOCTYPE html><html><head><style>' +
  'body{font-family:Google Sans,sans-serif;padding:20px;background:#f8f9fa;}' +
  'label{display:block;font-size:13px;font-weight:600;margin:12px 0 4px;}' +
  'input,select,textarea{width:100%;padding:8px;border:1px solid #dadce0;border-radius:4px;font-size:13px;}' +
  '.btn{background:#1a73e8;color:#fff;border:none;padding:10px 24px;border-radius:4px;cursor:pointer;font-size:14px;margin-top:16px;}' +
  '.btn:hover{background:#1557b0;}' +
  '#optionsGroup{display:none;}' +
  '</style></head><body>' +
  '<h3 style="margin:0 0 16px;color:#1a1a2e">Add New Column</h3>' +
  '<label>Column Name</label><input id="name" type="text" placeholder="e.g. Status"/>' +
  '<label>Column Type</label>' +
  '<select id="type" onchange="toggleOptions()">' +
  '<option value="TEXT">Text</option>' +
  '<option value="NUMBER">Number</option>' +
  '<option value="DATE">Date</option>' +
  '<option value="DROPDOWN">Dropdown (choose values)</option>' +
  '<option value="CHECKBOX">Checkbox</option>' +
  '<option value="CONTACT">Contact / Person</option>' +
  '<option value="PERCENT">Percentage</option>' +
  '<option value="DURATION">Duration (days)</option>' +
  '</select>' +
  '<div id="optionsGroup"><label>Dropdown Options <small>(one per line)</small></label>' +
  '<textarea id="options" rows="5" placeholder="In Progress&#10;Done&#10;Blocked"></textarea></div>' +
  '<button class="btn" onclick="submit()">Add Column</button>' +
  '<script>' +
  'function toggleOptions(){document.getElementById("optionsGroup").style.display=' +
  'document.getElementById("type").value==="DROPDOWN"?"block":"none";}' +
  'function submit(){' +
  'var name=document.getElementById("name").value.trim();' +
  'var type=document.getElementById("type").value;' +
  'var opts=document.getElementById("options").value.trim().split("\\n").filter(Boolean);' +
  'if(!name){alert("Enter a column name");return;}' +
  'google.script.run.withSuccessHandler(function(){google.script.host.close();}).createColumn(name,type,opts);}' +
  '</script></body></html>';
}

function getSetColumnTypeHTML(header, col) {
  return '<!DOCTYPE html><html><head><style>' +
  'body{font-family:Google Sans,sans-serif;padding:20px;}' +
  'label{display:block;font-size:13px;font-weight:600;margin:10px 0 4px;}' +
  'select{width:100%;padding:8px;border:1px solid #dadce0;border-radius:4px;}' +
  '.btn{background:#1a73e8;color:#fff;border:none;padding:10px 24px;border-radius:4px;cursor:pointer;margin-top:16px;}' +
  '</style></head><body>' +
  '<h3 style="margin:0 0 12px">Set Type: ' + header + '</h3>' +
  '<label>New Type</label>' +
  '<select id="type">' +
  ['TEXT','NUMBER','DATE','DROPDOWN','CHECKBOX','CONTACT','PERCENT','DURATION'].map(function(t){
    return '<option value="'+t+'">'+t+'</option>';
  }).join('') +
  '</select>' +
  '<button class="btn" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).createColumn(\''+header+'\',document.getElementById(\'type\').value,[])">Apply</button>' +
  '</body></html>';
}

function getManageDropdownHTML(header, col) {
  return '<!DOCTYPE html><html><head><style>' +
  'body{font-family:Google Sans,sans-serif;padding:20px;}' +
  'label{display:block;font-size:13px;font-weight:600;margin:10px 0 4px;}' +
  'textarea{width:100%;height:160px;padding:8px;border:1px solid #dadce0;border-radius:4px;font-size:13px;}' +
  '.btn{background:#1a73e8;color:#fff;border:none;padding:10px 24px;border-radius:4px;cursor:pointer;margin-top:12px;}' +
  '</style></head><body>' +
  '<h3 style="margin:0 0 12px">Edit Dropdown: ' + header + '</h3>' +
  '<label>Options <small>(one per line)</small></label>' +
  '<textarea id="opts" placeholder="In Progress&#10;Done&#10;Blocked"></textarea>' +
  '<button class="btn" onclick="' +
  'var opts=document.getElementById(\'opts\').value.trim().split(\'\\n\').filter(Boolean);' +
  'google.script.run.withSuccessHandler(function(){google.script.host.close();}).updateDropdownOptions(' + col + ',opts);">Save</button>' +
  '</body></html>';
}
