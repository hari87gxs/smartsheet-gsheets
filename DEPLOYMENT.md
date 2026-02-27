# ProjectSheet Pro — Deployment Guide

> Full step-by-step guide to deploy and run the Smartsheet-equivalent Google Sheets Add-on from scratch.

---

## Prerequisites

- A **Google account** (personal or Google Workspace)
- A **Google Sheet** (new or existing) where you want the add-on to live
- Access to **Google Apps Script** (script.google.com)

---

## Part 1 — Initial Setup (Do Once)

### Step 1 — Create or open your Google Sheet

1. Go to [sheets.new](https://sheets.new) to create a new spreadsheet, or open an existing one
2. Rename it (e.g. `My Project`)
3. Note the spreadsheet URL — you'll need it later

---

### Step 2 — Open the Apps Script editor

From your Google Sheet:

```
Extensions → Apps Script
```

This opens the script editor at `script.google.com`.

---

### Step 3 — Copy all script files from GitHub

Go to: **https://github.com/hari87gxs/smartsheet-gsheets**

For each file below, click the file on GitHub → click **Raw** → copy all content → paste into Apps Script:

#### Server-side `.gs` files

| GitHub file | Action in Apps Script |
|---|---|
| `Code.gs` | Click the default `Code.gs` file → select all → paste |
| `Setup.gs` | Click `+` (Add file) → Script → name it `Setup` → paste |
| `RowHierarchy.gs` | Add file → Script → name it `RowHierarchy` → paste |
| `ColumnTypes.gs` | Add file → Script → name it `ColumnTypes` → paste |
| `Automations.gs` | Add file → Script → name it `Automations` → paste |
| `ActivityLog.gs` | Add file → Script → name it `ActivityLog` → paste |
| `Collaboration.gs` | Add file → Script → name it `Collaboration` → paste |
| `Utils.gs` | Add file → Script → name it `Utils` → paste |

#### HTML files

| GitHub file | Action in Apps Script |
|---|---|
| `Gantt.html` | Add file → HTML → name it `Gantt` → paste |
| `Kanban.html` | Add file → HTML → name it `Kanban` → paste |
| `CalendarView.html` | Add file → HTML → name it `CalendarView` → paste |
| `Dashboard.html` | Add file → HTML → name it `Dashboard` → paste |
| `Automations.html` | Add file → HTML → name it **`Automationshtml`** (no dot, no space) → paste |
| `Sidebar.html` | Add file → HTML → name it `Sidebar` → paste |

> ⚠️ **Critical:** The Automations HTML file **must** be named `Automationshtml` (not `Automations`) in Apps Script, because Apps Script does not allow a `.gs` file and `.html` file to share the same base name.

---

### Step 4 — Update the manifest (`appsscript.json`)

In the Apps Script editor:

1. Click **Project Settings** (gear icon ⚙️ in the left sidebar)
2. Check **"Show `appsscript.json` manifest file in editor"**
3. Go back to the editor, click `appsscript.json`
4. Replace the entire content with:

```json
{
  "timeZone": "Asia/Singapore",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/script.send_mail",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

5. Press **Cmd+S** (Mac) / **Ctrl+S** (Windows) to save

> ⚠️ `script.container.ui` is **required** for `showModalDialog` (Gantt, Kanban, Calendar, Dashboard). Without it you get: `Exception: Specified permissions are not sufficient to call Ui.showModalDialog`.

---

### Step 5 — Save all files

Press **Cmd+S** in the Apps Script editor. All files must show no unsaved indicator.

---

### Step 6 — Run `installTriggers`

This installs the `onEdit` and `onOpen` triggers required for automations and activity logging.

1. In the Apps Script editor, click **`Utils.gs`** in the left panel
2. In the function dropdown at the top toolbar, select **`installTriggers`**
3. Click **▶ Run**
4. An authorization popup will appear — click **Review permissions** → choose your account → click **Advanced** → **Go to project (unsafe)** → **Allow**
5. After authorization, run `installTriggers` again — you should see: `✅ All triggers installed`

---

### Step 7 — Reload your Google Sheet

Go back to your Google Sheet tab and **refresh the page (Cmd+R / F5)**.

You should now see a **📋 ProjectSheet** menu in the top menu bar.

---

## Part 2 — First Use

### Step 8 — Initialize your project sheet

From the Google Sheet menu:

```
📋 ProjectSheet → Setup → 🏗 New Gantt Project (with sample data)
```

This creates:
- System columns A (`_indent`), B (`_id`), C (`_locked`) — hidden helpers
- Data columns: Task Name, Assigned To, Start Date, End Date, Duration, Status, % Complete, Priority, Predecessor, Notes
- 5 sample rows with hierarchy to get you started
- Hidden system sheets: `_PS_META`, `_PS_ACTIVITY`, `_PS_AUTOMATIONS`

> Alternatively use **🗒 Blank Project** if you want to start empty, or **📋 Kanban Project** for a kanban-only layout.

---

### Step 9 — Enter your data

Your data sheet has headers in **Row 1** and data from **Row 2** onwards.

The column layout is:

| Col | Header | Description |
|-----|--------|-------------|
| A | `_indent` | Row hierarchy level (0=parent, 1=child, 2=grandchild) — set via menu |
| B | `_id` | Unique row ID — auto-populated |
| C | `_locked` | Lock flag — set via menu |
| D | `Task Name` | Task title |
| E | `Assigned To` | Assignee name |
| F | `Start Date` | Date format: `MM/DD/YYYY` |
| G | `End Date` | Date format: `MM/DD/YYYY` |
| H | `Duration` | Number of days |
| I | `Status` | `Done`, `In Progress`, `To Do`, `Blocked`, `In Review`, `Cancelled` |
| J | `% Complete` | Number 0–100 |
| K | `Priority` | `Critical`, `High`, `Medium`, `Low` |
| L | `Predecessor` | Row number of predecessor task |
| M | `Notes` | Free text |

> ⚠️ **Date format:** Enter dates as `MM/DD/YYYY` (e.g. `02/01/2026`). After entering, select the date columns → **Format → Number → Date** to ensure Google Sheets treats them as real dates, not text.

---

## Part 3 — Using the Views

### Gantt Chart

```
📋 ProjectSheet → 📊 Gantt Chart
```

- **Day / Week / Month** zoom buttons at top
- **Today** button scrolls to current date
- **Critical Path** highlights tasks with predecessor chains
- **Collapse All / Expand All** controls hierarchy visibility
- Filter by **Status** or **Assignee** using the dropdowns
- Drag the splitter between task list and bars to resize

**Requires:** `Start Date` and `End Date` columns populated.

---

### Kanban Board

```
📋 ProjectSheet → 🃏 Kanban Board
```

- Tasks appear as cards in columns matching their `Status` value
- **Drag cards** between columns to update status live in the sheet
- Click **+ Add card** in any column to add a new task
- Filter by assignee or priority using toolbar dropdowns
- Click **+ Column** to add a custom status column

**Requires:** `Status` column populated with values like `To Do`, `In Progress`, `Done`, `Blocked`.

---

### Calendar View

```
📋 ProjectSheet → 📅 Calendar View
```

- **Month** and **Week** toggle at top
- Tasks appear on their `Start Date` and `End Date`
- Hover a task to see tooltip with full details
- Filter by assignee

**Requires:** `Start Date` or `End Date` columns populated.

---

### Dashboard

```
📋 ProjectSheet → 📈 Dashboard
```

Shows:
- **Total / Completed / In Progress / Overdue** stat cards
- **Status Breakdown** donut chart
- **Tasks by Assignee** bar chart
- **Priority Mix** bar chart
- **Upcoming Deadlines** (next 14 days)
- **Blocked Items** list
- **Weekly Completion Rate** velocity chart

**Requires:** `Status`, `Assigned To`, `Priority`, and `End Date` columns for full functionality.

---

### Automation Rules

```
📋 ProjectSheet → ⚡ Automations → Manage Rules
```

Create rules with:
- **Triggers:** Field Changed, Status Changed, Row Added, Date Reached
- **Actions:** Send Email, Change Status, Set Field, Add Comment, Lock Row, Call Webhook
- Template variables: `{{task_name}}`, `{{assigned_to}}`, `{{status}}`, `{{sheet_url}}`

---

### Activity Log

```
📋 ProjectSheet → 📝 Activity Log
```

Every cell edit is recorded with: timestamp, user, action, old value, new value. Stored in the hidden `_PS_ACTIVITY` sheet.

---

## Part 4 — Row Hierarchy

To create parent/child structure (like Smartsheet's hierarchy):

1. Click a row you want to make a **child** of the row above it
2. Go to: `📋 ProjectSheet → Hierarchy → ↘ Indent Row`
3. Repeat to add more levels (grandchild = indent twice)
4. To reverse: `📋 ProjectSheet → Hierarchy → ↖ Outdent Row`
5. To collapse children: `📋 ProjectSheet → Hierarchy → Collapse Children`

In the Gantt chart, parent rows appear **bold** with a thicker bar.

---

## Part 5 — Troubleshooting

### Views show "No tasks found"

**Cause:** `getRowTree()` can't find data in the active sheet.

**Fix checklist:**
1. Make sure your data sheet (e.g. `ECC` or `Sheet1`) is the **active tab** when you open a view
2. Confirm **Row 1 has headers** starting at column A (or D if using system columns)
3. Confirm **Row 2+ has data**
4. Open Apps Script → `Utils.gs` → run `debugGetRowTree` to see exactly what the server finds

---

### "Loading tasks…" spinner never goes away

**Cause:** `SpreadsheetApp.getActiveSheet()` was being called from client-side JavaScript (now fixed).

**Fix:** Make sure your HTML files contain `.getRowTree()` (no argument), **not** `.getRowTree(SpreadsheetApp.getActiveSheet())`.

---

### `Exception: Specified permissions are not sufficient to call Ui.showModalDialog`

**Cause:** `appsscript.json` is missing the `script.container.ui` scope.

**Fix:** Add `"https://www.googleapis.com/auth/script.container.ui"` to the `oauthScopes` array in `appsscript.json` → Save → run `installTriggers` again to trigger re-authorization.

---

### Authorization: "This app isn't verified"

**Cause:** Unverified script on a Google Workspace domain.

**Fix:** Click **Advanced** → **Go to project (unsafe)** → **Allow**. This is safe for your own scripts.

---

### Automations dialog won't open / file not found

**Cause:** The HTML file was named `Automations` instead of `Automationshtml`.

**Fix:** In Apps Script editor, rename the Automations HTML file to `Automationshtml`. Confirm that `Automations.gs` line 73 says `createHtmlOutputFromFile('Automationshtml')`.

---

### `getRowTree` returns 0 rows (seen in debug log)

**Cause:** Date cells in the sheet contain JavaScript `Date` objects which cannot be serialized by `google.script.run`.

**Fix:** Make sure `Utils.gs` contains the Date → ISO string conversion inside `getRowTree`:
```javascript
if (val instanceof Date) {
  val = isNaN(val.getTime()) ? '' : val.toISOString();
}
```
This is already in the current version. If you see this issue, re-paste `Utils.gs` from GitHub.

---

## Part 6 — Updating the Add-on

When new fixes are pushed to GitHub:

1. Go to **https://github.com/hari87gxs/smartsheet-gsheets**
2. For each changed file: click file → **Raw** → copy all
3. In Apps Script editor: click the matching file → select all → paste → **Save**
4. Reload your Google Sheet

> You do **not** need to re-run `installTriggers` for HTML-only changes.  
> Re-run `installTriggers` only if `appsscript.json` scopes or `.gs` trigger functions changed.

---

## File Reference

```
smartsheet-gsheets/
├── appsscript.json          ← Manifest + OAuth scopes (paste into Apps Script manifest)
│
├── Code.gs                  ← Menu builder + view launchers (onOpen, openGanttView, etc.)
├── Setup.gs                 ← Project templates (setupBlankProject, setupGanttProject, etc.)
├── RowHierarchy.gs          ← Row indent/outdent/collapse/expand
├── ColumnTypes.gs           ← Typed column management (dropdowns, dates, etc.)
├── Automations.gs           ← Rule engine (trigger + action evaluation)
├── ActivityLog.gs           ← Per-cell edit history stored in _PS_ACTIVITY sheet
├── Collaboration.gs         ← Row comments + @mention email + share dialog
├── Utils.gs                 ← getRowTree, updateTaskStatus, installTriggers, debug helpers
│
├── Gantt.html               ← Interactive Gantt chart modal  → Apps Script name: Gantt
├── Kanban.html              ← Drag-drop Kanban board modal   → Apps Script name: Kanban
├── CalendarView.html        ← Month/week calendar modal      → Apps Script name: CalendarView
├── Dashboard.html           ← Project metrics dashboard      → Apps Script name: Dashboard
├── Automations.html         ← Automation rule builder UI     → Apps Script name: Automationshtml ⚠️
└── Sidebar.html             ← Settings + quick-launch panel  → Apps Script name: Sidebar
```

---

## Quick Reference Card

| What you want to do | Where |
|---|---|
| Open Gantt chart | `📋 ProjectSheet → 📊 Gantt Chart` |
| Open Kanban board | `📋 ProjectSheet → 🃏 Kanban Board` |
| Open Calendar | `📋 ProjectSheet → 📅 Calendar View` |
| Open Dashboard | `📋 ProjectSheet → 📈 Dashboard` |
| Add a project template | `📋 ProjectSheet → Setup` |
| Indent a row (make child) | `📋 ProjectSheet → Hierarchy → ↘ Indent Row` |
| View edit history | `📋 ProjectSheet → 📝 Activity Log` |
| Create automation rule | `📋 ProjectSheet → ⚡ Automations → Manage Rules` |
| Re-install triggers | Run `installTriggers` from `Utils.gs` in Apps Script |
| Debug data loading | Run `debugGetRowTree` from `Utils.gs` in Apps Script |
| Share spreadsheet | `📋 ProjectSheet → 👥 Share` |
| Export as PDF | `📋 ProjectSheet → Export → 📄 Export PDF` |
