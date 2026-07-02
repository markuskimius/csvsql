# CSVSQL

A browser-based CSV database application. Open CSV, Excel, and compressed files as database tables, run SQL queries, edit data inline, and save — all in a multi-window interface with no server, build step, or internet connection required. Fully self-contained with all dependencies bundled locally.

**[Try the live version](https://app.cbreak.org/csvsql/)**

<!-- Screenshots and video demos — add your own assets to a screenshots/ directory -->
<!-- ![CSVSQL workspace](screenshots/workspace.png) -->

## Features

- **Multiple file formats** — CSV, TSV, PSV, Excel (.xlsx/.xls), Gzip (.csv.gz), and ZIP archives
- **SQL queries** — Full SQLite syntax from the built-in console, including joins, subqueries, aggregates, UNION, CASE, and REGEXP. Query results are queryable tables too
- **SQL syntax highlighting** — Keywords, strings, numbers, comments, and identifiers are color-coded in the SQL console and filter inputs
- **Inline editing** — Click a cell to select it; press i, F2, or Ctrl/Cmd+U to edit; or Ctrl/Cmd+click a cell to edit directly. Tab/Shift+Tab/Enter navigate between cells while staying in edit mode. Enter returns to the column where editing started. Up/Down arrow commits and exits edit mode. Long text scrolls within the cell to keep the cursor visible. Escape to revert
- **Sort and filter** — Click the sort badge (triangle icon) in column headers to sort (multi-column with Shift+click on sort badges). Filter with SQL WHERE expressions including REGEXP. Excel-style column autofilter dropdowns with searchable checkboxes for point-and-click filtering. Status chips in the status bar show active features. Click Sorted or Filtered to clear; click Linked or Formatted to suspend/resume
- **Multi-window workspace** — Draggable, resizable subwindows with edge snapping. Tile, Grid, or Cascade layouts. Tabbing and docking support for IDE-style split layouts
- **Row and column management** — Right-click a row number to insert below or delete. Right-click a column header to insert a column to the right or delete. Right-click the `#` corner cell to insert a row or column when the table is empty. Ctrl/Cmd+click a column header to rename inline (duplicate names are rejected with a red border). Full undo/redo for all structural operations. Resize any column by dragging column borders — including the row-number column (double-click to auto-fit based on all rows, not just visible ones). When multiple columns are selected, double-click auto-fits all selected columns and drag-resize sets all selected columns to the same width
- **Column Manager** — Open via Edit > Manage Columns (Ctrl/Cmd+Shift+M) or right-click a column header. Searchable vertical list of all columns with multi-select (Shift+click, Ctrl/Cmd+click). Drag to reorder within the list or drop onto table headers. Alt+Up/Down to move selected columns. Double-click a column to scroll the table to it. All reorders batch into a single undo entry when the manager closes. Esc closes from anywhere
- **Cell selection** — Click a cell to highlight its row, column, and row number. Click a column header to select the entire column. Move the selection with arrow keys or vim-style h/j/k/l, extend to a rectangle with Shift+arrow (or Shift+H/J/K/L), Shift+click, or click-and-drag. Click a row number to select the entire row (drag or Shift+click for multiple rows). Ctrl+A selects all cells; Esc deselects all. Clicking the `#` corner cell toggles between select all and select none. With no cell selected, any arrow key focuses the cell in the middle of the view. Ctrl+←/→ moves the header-selected column
- **Clipboard** — Cut (Ctrl+X), Copy (Ctrl+C), and Paste (Ctrl+V) work on selected cells. Data is copied as tab-separated values. When a plugin display transform is active, copy uses formatted values; when formatting is disabled, raw values are copied. Copying with Select All or row selection includes the header row. Copy and cut work from any focus context (column header, titlebar, etc.)
- **Undo / Redo** — Ctrl+Z undoes cell edits, paste, cut, row insert/delete, column insert/delete, column rename, column reorder, and column resize. Ctrl+Shift+Z redoes. Multi-cell paste and cut undo as a single step. Works from anywhere in the window — no cell focus required
- **SELECT INTO** — Create new tables from query results (`SELECT ... INTO tablename ...`)
- **CREATE TABLE** — New tables created via SQL auto-open as editable windows
- **Drag and drop** — Drop files directly onto the window to open them
- **Open from URL** — Load data files from any HTTP/HTTPS URL
- **Save** — Write directly back to the original file (Chrome/Edge) or download. Save As supports CSV, TSV, PSV, Excel, Gzip, and ZIP formats
- **Frozen row-number column** — The `#` column stays pinned on the left edge when scrolling horizontally, so you always know which row you're looking at. Column headers use opaque backgrounds so scrolling data never shows through — including when a column is selected
- **Virtual scrolling** — Handles large datasets efficiently
- **AI analysis** *(experimental)* — Natural language data analysis with automatic SQL query execution, inline charts, formatted tables, and PDF report generation. Supports WebLLM (in-browser), Ollama (local), Claude, and OpenAI

## Installation

### Option 1: pip (recommended)

```sh
pip install csvsql
csvsql
```

This starts a local server and opens CSVSQL in your browser. If `csvsql` conflicts with another command on your system, use `csvsqlw` instead — it's an identical alias.

```
csvsql --port 9000          # custom port (default: 8000)
csvsql --no-browser          # don't auto-open browser
csvsql --host 0.0.0.0       # bind to all interfaces
```

### Option 2: No install — open directly

All dependencies are bundled locally — no internet connection needed. Clone the repo and open `index.html` in your browser:

```sh
git clone https://github.com/markuskimius/csvsql.git
open csvsql/index.html
```

Or serve with any static file server:

```sh
python3 -m http.server 8000
# open http://localhost:8000
```

## Usage

### Opening Files

Use **File > Open** (Ctrl+O / Cmd+O), **File > Open URL**, or drag and drop files onto the window.

| Format | Extensions | Notes |
|--------|-----------|-------|
| CSV | .csv, .txt | Delimiter auto-detected (comma, tab, pipe, etc.) |
| TSV | .tsv | Tab-delimited |
| PSV | .psv | Pipe-delimited |
| Excel | .xlsx, .xls | Each non-empty worksheet opens as a separate table |
| Gzip | .csv.gz, etc. | Decompressed in browser; inner file opened by type |
| ZIP | .zip | All recognized data files inside the archive are opened |

<!-- ![Opening a file](screenshots/open-file.png) -->

### Editing

- **Edit cells** — Click a cell to select it; press i, F2, or Ctrl/Cmd+U to enter edit mode; or Ctrl/Cmd+click a cell to edit it directly. Long text scrolls within the cell to keep the cursor visible
- **Navigate in edit mode** — Tab/Shift+Tab to move between cells (stays in edit mode), Enter to save and move down to the column where editing started (stays in edit mode, exits on last row), Up/Down arrow to commit and move to the adjacent row (exits edit mode), Escape to revert the edit
- **Navigate in select mode** — Arrow keys or h/j/k/l to move the selection, Tab/Shift+Tab to move right/left, Enter to move down, Escape to clear the selection
- **Insert/delete rows** — Right-click a row number and choose "Insert Row Below" or "Delete Row"
- **Insert/delete columns** — Right-click a column header and choose "Insert Column Right" or "Delete Column"
- **Empty table entry point** — Right-click the `#` corner cell to insert a row or column when the table has no data
- **Rename columns** — Ctrl/Cmd+click a column header. Duplicate names are rejected with a red border on the input
- **Reorder columns** — Drag a column header to a new position. Click a column header to select it, then press Ctrl/Cmd+←/→ to nudge it
- **Column Manager** — Open from Edit > Manage Columns (Ctrl/Cmd+Shift+M) or right-click a column header. A searchable vertical list of all columns in a lightweight dialog. Click to select, Shift+click for range, Ctrl/Cmd+click to toggle. Drag to reorder within the list or drop onto the table's column headers. Alt+Up/Down moves selected columns. Double-click a column name to scroll the table to it. All reorders are batched into a single undo entry when the manager closes. Esc closes from anywhere
- **Resize columns** — Drag any column border to resize — the resize handle spans both sides of the divider line, including the `#` row-number column. Double-click to auto-fit the column to its content (measures all rows, not just visible ones). When multiple columns are selected (e.g. via Ctrl+A), double-click auto-fits all selected columns, and drag-resize sets all selected columns to the dragged column's width on release
- **Rename tables** — Ctrl/Cmd+click the window title
- **Highlight row & column** — Clicking a cell highlights its row and column. Click a column header to select the entire column (with header row included for copy). Move the selection with arrow keys or vim h/j/k/l; extend to a rectangle with Shift+arrow (or Shift+H/J/K/L), Shift+click, or click-and-drag. Click a row number to select an entire row (drag or Shift+click for ranges). Ctrl+A selects all; Esc deselects all. Clicking the `#` corner cell toggles between select all and select none
- **Cut / Copy / Paste** — Select cells and use Ctrl+X, Ctrl+C, Ctrl+V (Cmd on Mac). Data is copied as TSV. When a plugin display transform is active, copy uses formatted values; when formatting is disabled, raw values are copied. Select All and row selection copies include the header row. Works from any focus context (column header, titlebar, etc.)
- **Undo / Redo** — Ctrl+Z / Ctrl+Shift+Z (Cmd on Mac). Undoes cell edits, paste, cut, row insert/delete, column insert/delete, column rename, column reorder, and column resize. Multi-cell operations undo as a single step. Works from anywhere — no cell focus required

### Touch Gestures

- **Edit a cell** — Double-tap the cell
- **Select a rectangle of cells** — Tap a cell, then tap and pan to draw a selection rectangle from the first cell to the panned cell
- **Reorder a column** — Tap a column header, then tap-and-hold the same header and pan to the target position
- **Move a window** — Tap the title bar, then tap the title bar again and pan to drag the window

<!-- ![Inline editing](screenshots/editing.png) -->

### Sorting and Filtering

- **Sort** — Click the sort badge (triangle icon) in a column header to cycle: ascending → descending → unsorted
- **Multi-column sort** — Shift+click additional sort badges. Numbers inside triangle badges show sort priority
- **Filter** — Type a SQL WHERE expression in the filter bar (without the `WHERE` keyword):
  ```
  age > 30 AND name LIKE '%Smith%'
  name REGEXP 'smith|jones'
  ```
- **Column autofilter** — Click the funnel icon on any column header to open an Excel-style dropdown with searchable checkboxes for each unique value. Uncheck values to hide matching rows. Multiple column filters AND together and combine with the WHERE filter. The funnel fills teal when a filter is active. Click the **Filtered** chip in the status bar to clear all filters at once
- **Status chips** — **Sorted** and **Filtered** chips are always shown in the status bar center. **Linking** only appears when the table is a link source, **Linked** only when it is a link target, and **Formatted** only when a plugin has matching column transform rules. Inactive chips appear with faint outlines. **Sorted** and **Filtered** chips clear the sort or filters when clicked (chip becomes inactive). **Linked** (on target tables receiving link filters) and **Formatted** chips toggle suspend/resume — suspended features show the chip with strikethrough. A **Linking** chip (coral red) appears on the source table whose selection drives link filters and on intermediate tables in a transitive chain; click to suspend/resume outbound linking. When a table becomes a link target (receives link filters from another source), any previously suspended Linking state on that table is reset. Keyboard shortcuts: Ctrl/Cmd+Shift+1 (clear sort), 2 (clear filters), 3 (toggle link), 4 (toggle format)
- **Column badges** — The sort badge (orange triangle) and filter button (funnel icon) are rendered inline inside the column header as equal-sized clickable buttons. Both the sort triangle outline and funnel outline light up only when hovering their own button. Click the sort badge to sort; click the filter button to open the autofilter. The sort badge shows a filled triangle when sorted or a faint outline when unsorted. The filter button shows a faint funnel outline when inactive or a filled teal funnel when a filter is active. Small linking (coral red dot), link (green dot), and format (pink dot) badges appear centered over the bottom border of column headers — each badge type only renders when the table is relevant for that feature (link source, link target, or has matching transforms respectively). Active badges are filled, inactive badges show faint outlines. The linking badge appears on columns used for outbound linking; the link badge appears on columns receiving link filters. For transitive links, badges show a depth number (2 for secondary, 3 for tertiary, etc.). Badge colors match the status bar chips

<!-- ![Sorting and filtering](screenshots/filter.png) -->

### SQL Console

The SQL Console at the bottom runs queries against all open tables using SQLite syntax. Press **Ctrl+Enter** (Cmd+Enter on Mac) to execute. The console and filter inputs feature SQL syntax highlighting. Query results open as new queryable tables.

```sql
-- Query a loaded table by its filename (minus extension)
SELECT * FROM sample WHERE name LIKE 'A%'

-- Aggregate queries
SELECT department, AVG(salary) as avg_salary FROM employees GROUP BY department

-- Join across tables
SELECT a.name, b.value FROM table1 a JOIN table2 b ON a.id = b.id

-- REGEXP (case-insensitive)
SELECT * FROM employees WHERE name REGEXP '^(John|Jane)'

-- Create a new table from query results
SELECT name, salary INTO high_earners FROM employees WHERE salary > 100000

-- Create an empty table
CREATE TABLE projects (id, name, status)
```

Table names are derived from the filename (e.g., `employees.csv` → `employees`). All column values are stored as TEXT.

INSERT, UPDATE, DELETE, ALTER TABLE, and DROP TABLE all work. Changes to existing tables are reflected in their windows immediately.

<!-- ![SQL Console](screenshots/sql-console.png) -->

### AI Analysis *(experimental)*

The AI tab lets you analyze data using natural language. The AI automatically writes and executes SQL queries against your full dataset, so it works with tables of any size. You can also chat with the AI without any tables loaded.

**Providers:**

| Provider | Type | Setup |
|----------|------|-------|
| WebLLM (default) | In-browser | No setup — runs via WebGPU in Chrome/Edge 113+ |
| Ollama | Local | Install from [ollama.com](https://ollama.com), run `ollama pull llama3.2` |
| Claude | Cloud | API key from [console.anthropic.com](https://console.anthropic.com) |
| OpenAI | Cloud | API key from [platform.openai.com](https://platform.openai.com) |

Type a question and press **Enter** to send. Use **Shift+Enter** for multiline prompts and **Up/Down** arrows for prompt history. Click the gear icon to configure provider, model, and API keys.

The AI receives column statistics and sample rows, then writes SQL queries to get exact answers. Queries are executed automatically and results fed back for up to 5 rounds of analysis.

**Rich output:** The AI can render inline charts (Chart.js), formatted tables, and downloadable PDF reports. Ask for a visualization, a formatted table, or a PDF report and it will appear inline in the chat. Chart.js and jsPDF are loaded on demand when first needed. Drag and drop images (PNG, JPG) onto the AI chat area to upload them for inclusion in PDF reports (e.g., company logos).

### Saving Files

- **Save** (Ctrl+S / Cmd+S) — Writes directly back to the original file on Chrome/Edge (via File System Access API). On Firefox, triggers a download
- **Save As** — Prompts for a new filename. Supports CSV, TSV, PSV, Excel (.xlsx), Gzip, and ZIP formats
- **ZIP archives** — Saving any table from a ZIP re-packs all tables from that archive into the same ZIP
- **Excel workbooks** — Saving any sheet re-packs all sheets into the same workbook

### Window Management

- **Move** — Drag the title bar. Windows snap to workspace edges and other window edges within 10 px
- **Resize** — Drag any edge or corner. Edges snap to workspace boundaries and other window edges
- **Maximize/Restore** — Double-click the title bar, or click the maximize button
- **Minimize** — Click the minimize button. Restore from the Windows menu
- **Close** — Click the close button. Ctrl/Cmd+click closes all windows
- **Layout** — Use the Windows menu to Tile Horizontally, Tile Vertically, Grid, or Cascade
- **Proportional scaling** — Windows reposition and resize proportionally when the browser window or console panel is resized
- **Dialog boxes** — Prompts (New Table, Save As, Insert Column, Open URL) and AI Settings open as modal dialog boxes: the workspace dims behind them and they block other windows until you confirm, press Esc, or click outside to dismiss. All dialog boxes — including these, the Column Manager, and the About windows — are draggable and resizable but do not snap to window edges

### Tabbing and Docking

Windows can be combined into tabbed groups and split layouts for an IDE-style workspace.

- **Tab windows** — Hold Shift and drag a window onto another window's title bar to merge them into a tab group. Click tabs to switch between windows
- **Split dock** — Hold Shift and drag a window onto the body area of another window. Drop zones are divided diagonally — drop on the top, right, bottom, or left region to split that direction
- **Reorder tabs** — Drag a tab left or right within the tab bar to rearrange it
- **Move tabs** — Hold Shift and drag a tab to move it to another window or dock pane. A ghost preview shows where the window will land. Dropping on empty space moves the window to the drop location
- **Undock** — Hold Shift and drag a tab outside its dock container to detach it as a standalone window. The undocked window retains the pane's size
- **Splitter** — Drag the divider between split panes to resize. Double-click to reset to 50/50
- **Maximize** — Double-click a tab or the empty tab bar area to maximize/restore the dock container
- **Rename** — Ctrl/Cmd+click a tab to rename the table
- **Close tab** — Click the ✕ on a tab. When a dock reduces to a single tab, it dissolves back to a standalone window

<!-- ![Window layouts](screenshots/layouts.png) -->

## Plugins

Plugins customize how cell values are displayed. A plugin is a JSON file that maps table and column name patterns (regex) to display expressions in the CSVSQL expression language — a safe, sandboxed language with no JavaScript execution.

Multiple plugins can be loaded simultaneously and stack on the same table — each column is governed by the last-loaded plugin with a matching rule.

**Loading:** Use **Plugins > Load Plugin** in the menu, or drag and drop a `.json` file onto the app. A toast notification confirms success or shows errors. Plugins persist across page reloads.

**Plugin management:** Loaded plugins appear in the Plugins menu with an ✕ button for quick unloading — the menu stays open so you can unload multiple plugins without reopening it. Click a plugin's name to open an About dialog showing version, author, creation date, description, matching rules, and an Unload button.

**Toggle:** Columns with an active transform show a pink dot badge (format badge) at the bottom-center of the header. The **Formatted** chip in the status bar toggles all transforms on/off (Ctrl/Cmd+Shift+4). A **Sorted** chip appears when sort is active (click to clear), and a **Filtered** chip when filters are active (click to clear).

**Autofilter integration:** When a transform is active, the column's autofilter dropdown shows formatted display values and searches against them.

**Plugin formats:** Two formats are supported. The legacy format has `table` and `columns` at the top level. The multi-table format uses a `tables` array to target different tables with different rules in a single plugin:

```json
{
  "name": "Orders ↔ Customers ↔ Products",
  "version": "1.1.0",
  "author": "CSVSQL",
  "created": "2025-06-13",
  "description": "Cross-table linking with display formatting",
  "tables": [
    {
      "table": "orders",
      "columns": [
        { "match": "^total$", "display": "isNum(value) ? '$' + fixed(num(value), 2) : value" },
        { "match": "^order_date$", "display": "isEmpty(value) ? '' : date(value, 'locale')" }
      ]
    },
    {
      "table": "customers",
      "columns": [
        { "match": "^signup_date$", "display": "isEmpty(value) ? '' : date(value, 'locale')" }
      ]
    }
  ],
  "links": [
    {
      "source": { "table": "customers", "column": "^id$" },
      "target": { "table": "orders", "column": "^customer_id$" }
    },
    {
      "source": { "table": "orders", "column": "^customer_id$" },
      "target": { "table": "customers", "column": "^id$" }
    }
  ]
}
```

The `version`, `author`, `created`, and `description` fields are optional metadata displayed in the About dialog. The `tables` and `links` arrays are both optional — a plugin can have just display rules, just links, or both.

**Cross-table linking:** The `links` array defines relationships between tables. When you select rows in a source table, the target table is automatically filtered to matching values. All table and column patterns are regex. The source table is excluded from its own link targets (even with `.*`). Link-filtered columns show a green dot badge (link badge) at the bottom-center of the header and a **Linked** chip on target tables. Outbound linking columns show a coral red dot badge (linking badge). A **Linking** chip appears on the source table and on intermediate tables in a transitive chain; click to suspend/resume outbound linking — the chip stays visible when suspended. For transitive links, depth numbers appear inside the linking and link badges (2 for secondary, 3 for tertiary, etc.). When a table receives link filters from a new source, any previously suspended Linking state on that table is reset. Clearing the selection clears the link filter. Link filters are separate from manual column autofilters. Links propagate transitively — selecting a product filters orders for that product, which in turn filters customers who placed those orders. Cycle detection and a depth limit prevent infinite loops.

Bundled example plugins and sample CSV files are in the `example/` directory: date formatting, USD currency, boolean display, ID zero-padding, and a linked-tables demo with orders, customers, and products.

The in-app **Plugins > Expression Reference** has the full language documentation including all operators, built-in functions, and examples.

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+O / Cmd+O | Open file |
| Ctrl+S / Cmd+S | Save table |
| Ctrl+N / Cmd+N | New table |
| Ctrl+W / Cmd+W | Close window |
| Ctrl+← / Ctrl+→ (or Cmd+arrow on Mac) | Move header-selected column left / right |
| i, F2, or Ctrl+U / Cmd+U | Enter edit mode on the selected cell |
| Ctrl+click / Cmd+click | Enter edit mode on the clicked cell directly |
| / (cell selected, not editing) | Jump to the window's filter input |
| Escape (in filter input) | Return focus to the selected cell |
| Tab / Shift+Tab (cell selected, not editing) | Move selection right / left |
| Enter (cell selected, not editing) | Move selection down |
| Ctrl+Shift+L / Ctrl+Shift+H (cell selected, not editing) | Switch to next / previous table window |
| Ctrl+H / J / K / L (cell selected, not editing) | Nudge the active window 5 px left / down / up / right |
| Arrow keys (no cell selected) | Focus the cell in the middle of the view |
| Arrow keys or h/j/k/l (cell selected, not editing) | Move selection to the adjacent cell |
| Up / Down arrow (editing) | Commit edit and move to adjacent row (exits edit mode) |
| Shift+arrow or Shift+H/J/K/L | Extend cell selection — highlights each selected cell's row & column |
| Ctrl+A / Cmd+A | Select all cells |
| Esc (cell selected, not editing) | Deselect all cells |
| Ctrl+X / Cmd+X | Cut selected cells |
| Ctrl+C / Cmd+C | Copy selected cells |
| Ctrl+V / Cmd+V | Paste at selected cell |
| Ctrl+Z / Cmd+Z | Undo |
| Ctrl+Shift+Z / Cmd+Shift+Z | Redo |
| Ctrl+Shift+M / Cmd+Shift+M | Open Column Manager |
| Ctrl+Shift+1 / Cmd+Shift+1 through 4 | Clear Sort / Clear Filters / Toggle Link / Toggle Format |
| Ctrl+Enter / Cmd+Enter | Execute SQL query |
| Enter (AI tab) | Send AI prompt |
| Shift+Enter (AI tab) | Newline in AI prompt |
| Up / Down (AI tab) | AI prompt history |
| Tab / Shift+Tab (editing) | Move to next / previous cell (stays in edit mode) |
| Enter (editing) | Move down to the column where editing started (stays in edit mode) |
| Escape | Cancel cell edit |

## License

MIT — see [About dialog](javascript:void(0)) in the app for full text.
