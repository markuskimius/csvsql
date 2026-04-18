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
- Cell selection is a rectangle defined by `win.anchorCell` (fixed corner) and the focused cell (lead). Selected cells are cached as a Set of `"rownum:colName"` keys in `win.selectedCells`. Keys use column *names* + `_rownum` so the selection survives sort, filter, column reorder, and rename. `applyCellHighlights()` rebuilds `.row-highlight`/`.col-highlight`/`.cell-selected` classes from this Set after every `renderVisibleRows()` (needed for virtual scrolling). Anchor is reset by user-initiated focusin on a cell; programmatic focus (e.g., drag-end, arrow-key move) sets `win._programmaticFocus` to skip the reset
- Cell editability is a two-state model. Data cells render as `<td class="data-cell" tabindex="0">` — focusable for selection but **not** editable. `enterEditMode(td)` sets `contenteditable="true"` and places the caret at the end; `exitEditMode(td)` removes the attribute. F2, Ctrl/Cmd+U, and touch long-press are the entry points. The blur handler saves + exits edit mode; Escape either reverts the edit (in edit mode) or clears selection (in select mode). Keep the `data-cell` class as the identity check — don't rely on the presence of `contenteditable` to identify data cells
- Arrow-key behavior on a data cell depends on mode: in select mode, plain arrow collapses the selection to the adjacent single cell (`moveSingleCellSelection`), Shift+arrow extends the rectangle (`extendCellSelection`), and Ctrl/Cmd+←/→ moves the selection's columns as a block (`moveSelectionColumns`). In edit mode none of these fire — arrows pass through to native contenteditable caret movement
- Column reorder: plain drag on a header reorders (5 px threshold); Ctrl/Cmd+click renames; Ctrl/Cmd+←/→ nudges the header-selected column (`win.selectedCol`) when focus is outside a data cell, or the range of columns spanned by the cell selection when a data cell has focus. All paths funnel through `reorderColumn(win, fromIdx, toIdx)`, which splices `tables[name].columns`, re-registers the SQLite table, and calls `rebuildTable(win)`
- Touch gestures are layered on top of the mouse handlers. State lives at IIFE scope: `_touchLongPress` (cell long-press → edit mode after `LONG_PRESS_MS` = 600 ms), `_touchHeaderDrag` (active header drag), `_touchCellDrag` (active cell multi-select drag), `_touchWinDrag` (active titlebar window drag), and the 1.5-tap trackers `_lastHeaderTap` / `_lastCellTap` / `_lastTitleTap` (each keyed by a stable identifier — e.g. table + column name, window id + rownum + col — because `rebuildTable` recreates DOM nodes between taps). Cell long-press listens on both touch and pointer events because some browsers/devtools modes deliver one family but not the other; the handlers share `_touchLongPress` so whichever fires first claims the interaction and the other no-ops (pointer handlers filter out `pointerType === 'mouse'`, and additionally bail out when `_touchCellDrag` is active so the pointer path doesn't start a conflicting long-press mid-drag). The header, cell-select, and window-move 1.5-taps use touch events only. The cell 1.5-tap seeds `win.anchorCell` from the first-tap cell before starting the drag (real devices get this for free via the synthesized click → focusin, but we set it explicitly for robustness); `endCellDrag` calls `focusCellAt` (which uses `_programmaticFocus`) and preventDefaults touchend so the synthesized click doesn't collapse the selection back to a single cell. Data cells set `user-select: none` + `touch-action: manipulation` so iOS/Android don't hijack a still touch for text selection before the timer fires; edit mode overrides back to `user-select: text`. For edit-mode entry, the long-press timer flips `contenteditable="true"` synchronously and does NOT call `focus()` — the browser's touchend-synthesized click then lands on an already-editable cell, which iOS recognizes as a user gesture and opens the virtual keyboard natively. The header drag sets `th._didDrag = true` so the trailing click doesn't sort
