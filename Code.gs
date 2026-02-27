// ─────────────────────────────────────────────────────────────────────────────
// Code.gs — Main entry point
// ProjectSheet Pro: A full Smartsheet equivalent built on Google Sheets
// ─────────────────────────────────────────────────────────────────────────────

// ── Menu ──────────────────────────────────────────────────────────────────────
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('📋 ProjectSheet')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🆕 New Project')
      .addItem('Blank Project', 'setupBlankProject')
      .addItem('Project with Gantt Template', 'setupGanttProject')
      .addItem('Kanban Board Template', 'setupKanbanProject'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📊 Views')
      .addItem('📊 Gantt Chart', 'openGanttView')
      .addItem('🃏 Kanban Board', 'openKanbanView')
      .addItem('📅 Calendar View', 'openCalendarView')
      .addItem('🗂️ Dashboard', 'openDashboard'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️ Columns')
      .addItem('Add Column with Type…', 'addTypedColumn')
      .addItem('Set Column Type…', 'setColumnType')
      .addItem('Manage Dropdowns…', 'manageDropdowns'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🌿 Hierarchy')
      .addItem('Indent Row (child)', 'indentRow')
      .addItem('Outdent Row (parent)', 'outdentRow')
      .addItem('Collapse Children', 'collapseChildren')
      .addItem('Expand Children', 'expandChildren'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🤖 Automations')
      .addItem('Create Automation Rule…', 'openAutomationBuilder')
      .addItem('View All Rules', 'openAutomationList')
      .addItem('Run Automations Now', 'runAllAutomations'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📢 Collaboration')
      .addItem('Add Row Comment', 'addRowComment')
      .addItem('View Activity Log', 'openActivityLog')
      .addItem('Share & Notify…', 'openShareDialog'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📤 Export')
      .addItem('Export as PDF', 'exportPDF')
      .addItem('Export Gantt as PDF', 'exportGanttPDF')
      .addItem('Save Baseline Snapshot', 'saveBaseline'))
    .addItem('⚙️ Settings', 'openSettings')
    .addToUi();
}

// Card service homepage (Add-on sidebar)
function onHomepage(e) {
  return buildHomepageCard();
}

function onSheetsHomepage(e) {
  return buildHomepageCard();
}

function onFileScopeGranted(e) {
  return buildHomepageCard();
}

// ── View launchers ────────────────────────────────────────────────────────────
function openGanttView() {
  var html = HtmlService.createHtmlOutputFromFile('Gantt')
    .setWidth(1100).setHeight(650).setTitle('Gantt Chart');
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Gantt Chart');
}

function openKanbanView() {
  var html = HtmlService.createHtmlOutputFromFile('Kanban')
    .setWidth(1100).setHeight(650).setTitle('Kanban Board');
  SpreadsheetApp.getUi().showModalDialog(html, '🃏 Kanban Board');
}

function openCalendarView() {
  var html = HtmlService.createHtmlOutputFromFile('CalendarView')
    .setWidth(900).setHeight(650).setTitle('Calendar View');
  SpreadsheetApp.getUi().showModalDialog(html, '📅 Calendar View');
}

function openDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1100).setHeight(700).setTitle('Project Dashboard');
  SpreadsheetApp.getUi().showModalDialog(html, '🗂️ Project Dashboard');
}

function openSettings() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Homepage Card (Add-on sidebar) ────────────────────────────────────────────
function buildHomepageCard() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var meta = getProjectMeta(sheet);

  var card = CardService.newCardBuilder()
    .setName('ProjectSheet Pro')
    .setHeader(CardService.newCardHeader()
      .setTitle('ProjectSheet Pro')
      .setSubtitle(meta.name || 'Select or create a project')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/2x/table_chart_black_48dp.png'));

  // Stats section
  if (meta.exists) {
    var stats = getProjectStats(sheet);
    var statsSection = CardService.newCardSection()
      .setHeader('📊 Current Project Stats');
    statsSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('Total Tasks').setText(String(stats.total)));
    statsSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('Completed').setText(stats.done + ' (' + stats.pct + '%)'));
    statsSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('In Progress').setText(String(stats.inProgress)));
    statsSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('Blocked').setText(String(stats.blocked)));
    card.addSection(statsSection);
  }

  // Actions
  var actionsSection = CardService.newCardSection().setHeader('🚀 Views');
  actionsSection.addWidget(CardService.newTextButton()
    .setText('📊 Open Gantt Chart')
    .setOnClickAction(CardService.newAction().setFunctionName('openGanttView')));
  actionsSection.addWidget(CardService.newTextButton()
    .setText('🃏 Open Kanban Board')
    .setOnClickAction(CardService.newAction().setFunctionName('openKanbanView')));
  actionsSection.addWidget(CardService.newTextButton()
    .setText('📅 Calendar View')
    .setOnClickAction(CardService.newAction().setFunctionName('openCalendarView')));
  actionsSection.addWidget(CardService.newTextButton()
    .setText('🗂️ Dashboard')
    .setOnClickAction(CardService.newAction().setFunctionName('openDashboard')));
  card.addSection(actionsSection);

  // Quick add
  var addSection = CardService.newCardSection().setHeader('➕ Quick Actions');
  addSection.addWidget(CardService.newTextButton()
    .setText('Add Row Below')
    .setOnClickAction(CardService.newAction().setFunctionName('quickAddRow')));
  addSection.addWidget(CardService.newTextButton()
    .setText('Set Up New Project')
    .setOnClickAction(CardService.newAction().setFunctionName('setupBlankProject')));
  card.addSection(addSection);

  return card.build();
}

// ── Quick actions (card service callbacks) ────────────────────────────────────
function quickAddRow(e) {
  addRowBelow();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Row added!'))
    .build();
}

function addRowBelow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  sheet.insertRowAfter(row);
  logActivity('ROW_ADD', 'Row inserted at ' + (row + 1));
}
