# CSVSQL

A browser-based CSV database application. Open CSV, Excel, and compressed files as database tables, run SQL queries, edit data inline, and save — all in a multi-window interface with no server or build step required.

<!-- Screenshots and video demos — add your own assets to a screenshots/ directory -->
<!-- ![CSVSQL workspace](screenshots/workspace.png) -->

## Features

- **Multiple file formats** — CSV, TSV, PSV, Excel (.xlsx/.xls), Gzip (.csv.gz), and ZIP archives
- **SQL queries** — Full SQLite syntax from the built-in console, including joins, subqueries, aggregates, UNION, CASE, and REGEXP. Query results are queryable tables too
- **SQL syntax highlighting** — Keywords, strings, numbers, comments, and identifiers are color-coded in the SQL console and filter inputs
- **Inline editing** — Click any cell to edit. Tab/Enter to navigate, Escape to cancel
- **Sort and filter** — Click column headers to sort (multi-column with Shift+click). Filter with SQL WHERE expressions including REGEXP
- **Multi-window workspace** — Draggable, resizable subwindows. Tile, Grid, or Cascade layouts
- **Row and column management** — Add/delete rows, insert at position (right-click), add/rename/reorder columns
- **SELECT INTO** — Create new tables from query results (`SELECT ... INTO tablename ...`)
- **CREATE TABLE** — New tables created via SQL auto-open as editable windows
- **Drag and drop** — Drop files directly onto the window to open them
- **Open from URL** — Load data files from any HTTP/HTTPS URL
- **Save** — Write directly back to the original file (Chrome/Edge) or download. Save As supports CSV, TSV, PSV, Excel, Gzip, and ZIP formats
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

No dependencies needed. Clone the repo and open `index.html` in your browser:

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

- **Edit cells** — Click any cell to edit
- **Navigate** — Tab/Shift+Tab between cells, Enter to move down, Escape to cancel
- **Add rows** — Click `+ Row` in the toolbar, or right-click a row number to insert above
- **Delete rows** — Right-click a row number and choose Delete Row
- **Add columns** — Click `+ Col` in the toolbar
- **Rename columns** — Ctrl/Cmd+click a column header
- **Reorder columns** — Ctrl/Cmd+drag a column header to a new position
- **Rename tables** — Ctrl/Cmd+click the window title

<!-- ![Inline editing](screenshots/editing.png) -->

### Sorting and Filtering

- **Sort** — Click a column header to cycle: ascending → descending → unsorted
- **Multi-column sort** — Shift+click additional headers. Numbers next to arrows show sort priority
- **Filter** — Type a SQL WHERE expression in the filter bar (without the `WHERE` keyword):
  ```
  age > 30 AND name LIKE '%Smith%'
  name REGEXP 'smith|jones'
  ```

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

- **Move** — Drag the title bar
- **Resize** — Drag any edge or corner
- **Maximize/Restore** — Double-click the title bar, or click the maximize button
- **Minimize** — Click the minimize button. Restore from the Windows menu
- **Close** — Click the close button. Ctrl/Cmd+click closes all windows
- **Layout** — Use the View menu to Tile Horizontally, Tile Vertically, Grid, or Cascade
- **Proportional scaling** — Windows reposition and resize proportionally when the browser window or console panel is resized

<!-- ![Window layouts](screenshots/layouts.png) -->

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+O / Cmd+O | Open file |
| Ctrl+S / Cmd+S | Save table |
| Ctrl+N / Cmd+N | New table |
| Ctrl+W / Cmd+W | Close window |
| Ctrl+Enter / Cmd+Enter | Execute SQL query |
| Enter (AI tab) | Send AI prompt |
| Shift+Enter (AI tab) | Newline in AI prompt |
| Up / Down (AI tab) | AI prompt history |
| Tab / Shift+Tab | Navigate between cells |
| Enter | Move to next row |
| Escape | Cancel cell edit |

## License

MIT — see [About dialog](javascript:void(0)) in the app for full text.
