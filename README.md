# ProjectSheet Pro 📋

A full **Smartsheet equivalent** built as a Google Workspace Add-on / Apps Script project for Google Sheets.

## Features

| Feature | Status |
|---------|--------|
| 📊 Interactive Gantt Chart | ✅ Full (zoom, dependencies, critical path, filters) |
| 🃏 Kanban Board | ✅ Full (drag-drop, add cards, status columns) |
| 📅 Calendar View | ✅ Full (month/week, multi-field dates) |
| 📈 Dashboard | ✅ Full (donut chart, bar charts, deadlines, velocity) |
| 🌿 Row Hierarchy | ✅ Indent/outdent, collapse/expand |
| ⚙️ Column Types | ✅ Text, Dropdown, Date, Number, Checkbox, Formula |
| ⚡ Automations | ✅ Rule engine (field change, status change, date, row added) |
| 📋 Activity Log | ✅ Every change tracked with user + timestamp |
| 💬 Collaboration | ✅ Row comments, @mentions, email notifications |
| 📤 Export | ✅ PDF export, Baseline snapshots |
| ⚙️ Settings Sidebar | ✅ Stats, quick actions, preferences |

## File Structure

```
smartsheet-gsheets/
├── appsscript.json        # Add-on manifest & OAuth scopes
├── Code.gs                # Menu, view launchers, Card Service homepage
├── Setup.gs               # Project templates & column initialization
├── RowHierarchy.gs        # Indent/outdent, collapse/expand parent-child rows
├── ColumnTypes.gs         # Typed column management with UI dialogs
├── Automations.gs         # Automation rule engine + time trigger
├── ActivityLog.gs         # Change tracking (every edit logged)
├── Collaboration.gs       # Row comments, @mentions, sharing
├── Utils.gs               # Server-side helpers used by all HTML views
│
├── Gantt.html             # Interactive Gantt chart (1100×650)
├── Kanban.html            # Kanban board with drag-drop
├── CalendarView.html      # Month/week calendar
├── Dashboard.html         # Project metrics & charts
├── Automations.html       # Automation rules builder UI
├── Sidebar.html           # Settings sidebar
└── README.md
```

## Quick Start

### Option A — Direct Apps Script (Recommended)

1. Open [Google Sheets](https://sheets.google.com) → create a new spreadsheet
2. Go to **Extensions → Apps Script**
3. Delete the default `Code.gs` content
4. Copy each `.gs` file into the Apps Script editor (one file each)
5. Copy each `.html` file as new HTML files in the editor
6. Copy `appsscript.json` into the manifest (View → Show manifest file)
7. Save and reload the spreadsheet
8. Run **📋 ProjectSheet → Setup → New Blank Project** from the menu

### Option B — clasp (CLI deployment)

```bash
# Install clasp
npm install -g @google/clasp

# Login
clasp login

# Create new Apps Script project
clasp create --title "ProjectSheet Pro" --type sheets

# Push all files
clasp push

# Open the script editor
clasp open
```

### Option C — Install as Add-on

After pushing via clasp:
1. Go to Apps Script → Deploy → Test Deployments
2. Install as a Workspace Add-on
3. Open any Google Sheet — the sidebar and menu will appear

## Usage

### Creating a Project

Use **📋 ProjectSheet → New Project** to choose:
- **Blank Project** — empty grid with 11 standard columns
- **Gantt Template** — pre-filled with sample hierarchy + dates
- **Kanban Template** — board-optimised columns

### Views

| View | How to open |
|------|-------------|
| Gantt | Menu → Views → Gantt Chart |
| Kanban | Menu → Views → Kanban Board |
| Calendar | Menu → Views → Calendar View |
| Dashboard | Menu → Views → Dashboard |

### Automations

Menu → Automations → Create Automation Rule…

**Available Triggers:**
- When a field changes (any column + operator + value)
- When Status changes (specific value)
- When a row is added
- When a date is reached (on, before, or after)

**Available Actions:**
- Send an email (with `{{template_vars}}`)
- Change a field value
- Set another field
- Add a note/comment
- Lock the row
- Call a webhook

### Row Hierarchy

Select a row and use **Menu → Hierarchy** to indent/outdent.
Parent rows are bold; children are visually indented.
Collapse/expand children with one click.

### Comments

Select a row → **Menu → Collaboration → Add Row Comment**
Use `@name` in comments to send email notifications.

### Activity Log

Every edit is logged automatically (if triggers are installed).
View: **Menu → Collaboration → View Activity Log**

## Installing Triggers

Run **Utils.gs → `installTriggers()`** once to enable:
- `onEdit` — logs changes + fires automations
- `onOpen` — bootstraps system sheets
- Hourly — runs date-based automations

## Column Types

| Type | Behaviour |
|------|-----------|
| Text | Free text |
| Dropdown | Colour-coded single-select |
| Multi-select | Comma-separated values |
| Date | Date picker with validation |
| Number | Numeric with validation |
| Currency | Number with `$` format |
| Checkbox | TRUE/FALSE toggle |
| Formula | Custom formula |
| Auto-number | Sequential ID |
| Contact | Email field |

## License

MIT — free for personal and commercial use.
