# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

CSVSQL is a browser-based CSV database application. It treats CSV files as database tables with SQL query support, editable cells, and a multi-window interface. No build step or server required — open `index.html` directly in a browser. Also distributed via PyPI (`pip install csvsql`).

## Architecture

Single-page app with three core files:

- **index.html** — Shell: menubar, workspace area, SQL console panel, hidden file input, CDN script tags
- **style.css** — Dark theme styling, window management visuals, table layout
- **app.js** — All application logic in a single IIFE (`app` module), exposing methods on the global `app` object

### Key dependencies (bundled in `lib/`)

All dependencies are self-contained in the `lib/` directory — no CDN or internet required at runtime.

- **Papa Parse** — CSV parsing/unparsing
- **sql.js** — SQLite compiled to WebAssembly; each opened CSV is registered as a SQLite table
- **SheetJS (XLSX)** — Excel file reading/writing
- **JSZip** — ZIP archive support
- **Chart.js** — Inline chart rendering (lazy-loaded on first AI chart)
- **jsPDF + autotable** — PDF report generation (lazy-loaded on first AI PDF)
- **@mlc-ai/web-llm** — In-browser LLM inference via WebGPU (lazy-loaded)

### Data flow

1. CSV opened via File menu → Papa Parse parses → stored in `tables[name]` object (columns, rows, filename, modified flag)
2. Table registered in SQLite via `CREATE TABLE` + batch `INSERT` with prepared statements
3. Edits to cells update `tables[name].rows` and sync back to SQLite via `syncToSQL()`
4. SQL queries run against SQLite; results become new entries in `tables` and open in new subwindows
5. Save serializes `tables[name]` back to CSV via Papa.unparse and triggers browser download

### Window management

Custom subwindow system — each table/query result gets a draggable, resizable, minimizable window inside `#window-area`. Layout functions (tile, grid, cascade) reposition all visible windows. Windows reposition/resize proportionally when the browser window or console panel is resized (`scaleWindowsToArea()`). Windows track their own sort/filter state. The `windows` array and `tables` object are the two central data structures. Help windows (Manual, About) are singletons — reopening focuses the existing window instead of creating a duplicate.

### Row identity

Each row's primary key is its `_rownum` property (1-based index). Renumbered on insert/delete. Not a real column — excluded from CSV export.

## Development

No build tools. To develop:

```
python3 -m http.server 8000
# Then open http://localhost:8000
```

Or just open `index.html` directly — all dependencies are bundled locally.

## Testing

Playwright-based test suite (Chromium only). Tests are organized into `test/unit/`, `test/integration/`, and `test/e2e/`.

```bash
# Run all tests (installs deps + browser if needed)
./run-tests.sh

# Run all tests via npm
npm test

# Run a single test file
npx playwright test --config test/playwright.config.js test/integration/sql-queries.spec.js

# Run tests matching a grep pattern
npx playwright test --config test/playwright.config.js -g "filter"
```

The test server runs on port 8274 (auto-started by Playwright config). Tests use `?test=1` query param which sets `window._appReady` after init. Test helpers are in `test/helpers.js` — notably `openApp()`, `uploadFile()`, `executeSQL()`, `waitForWindow()`, and `getTableData()`.

Test fixtures live in `test/` (e.g., `sample1.csv`, `sample.xlsx`, `sample.zip`).

## PyPI Packaging

Published as `csvsql` on PyPI. The Python package in `csvsql/` serves static files via `cli.py`.

```bash
# IMPORTANT: Copy root static files + lib/ to csvsql/static/ before building
cp app.js index.html style.css csvsql/static/ && cp lib/* csvsql/static/lib/
rm -rf dist build *.egg-info && python3 -m build
python3 -m twine upload dist/*
```

Version is in `pyproject.toml`.

## Conventions

- All state lives in the `app` IIFE's closure (`windows`, `tables`, `nextWinId`, etc.)
- Public methods are returned from the IIFE and called from HTML onclick handlers or internally
- Test internals are exposed via `app._test` (only used by Playwright tests)
- Table and column names are sanitized to `[a-zA-Z0-9_]` for SQL compatibility
- SQL identifiers use bracket-quoting (`[tableName]`) to handle edge cases
- SQL syntax highlighting uses the overlay technique: a div with highlighted spans behind a transparent textarea/input. Tokenizer is `sqlHighlightHTML()`, setup is `setupSQLHighlight()` for the console and inline in `renderTableView()` for filter inputs
- Query result tables are registered in SQLite via `registerTable()` so they can be queried and filtered like any other table
- `db.export()` in sql.js destroys custom functions — `registerDBFunctions()` re-registers them after each export
- SELECT INTO is intercepted and handled manually (SQLite doesn't support it natively)
- AI analysis (experimental) uses a SQL tool-use loop: the AI writes SQL in ```sql code blocks, which are executed against SQLite and results fed back (up to 5 rounds). AI can also be used without any tables loaded (general chat mode)
- AI rich output: ```chart (Chart.js config JSON), ```table (columns/rows JSON), ```pdf (document spec JSON) blocks are post-processed into inline charts, HTML tables, and PDF download links. Chart.js and jsPDF are lazy-loaded via CDN on first use. PDFs support text, heading, table, chart, and image content blocks
- AI images: users can drag-and-drop PNG/JPG images onto the AI chat area. Stored in `_aiImages` (name → data URL). Available for inclusion in PDF reports via `{"type":"image","name":"filename.png"}`. Image drop on AI panel is intercepted before the global file-open drop handler
- AI providers: WebLLM (default, in-browser via WebGPU), Ollama (local), Claude (cloud), OpenAI (cloud)
- AI settings (provider, model, API keys) are persisted in localStorage under `csvsql_ai_settings`
- AI conversation history is kept in-memory (`_aiConversation` array) and cleared with the console
- AI prompt history (Up/Down arrow) is in-memory only, not persisted across sessions
- Console tab switching auto-focuses the corresponding input field (SQL input or AI prompt)
- Help windows (Manual, About) are singletons — `showHelpWindow()` focuses/restores existing window instead of creating a duplicate
