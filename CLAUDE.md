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

Custom subwindow system — each table/query result gets a draggable, resizable, minimizable window inside `#window-area`. Windows are constrained to the workspace area — drag, resize, and keyboard nudge all clamp position and size so windows cannot extend beyond the `#window-area` boundaries. All mouse-driven interactions (window drag, window resize, dock resize, tab drag, splitter drag, column resize, console resize, cell/row selection) require left mouse button only (`e.button !== 0` guard). Right/middle click on data cells calls `e.preventDefault()` to suppress focus and prevent accidental selection. During drag and resize, windows snap to workspace edges and other window edges within `SNAP_THRESHOLD` (10 px). Snap helpers: `getSnapEdges(excludeWin)` collects edge positions from the area and all visible windows, `findSnap(value, edges, threshold)` finds the nearest edge, `snapPosition()` applies snap to both axes during drag, and the resize handler snaps each active edge independently. Blue guide lines (`showSnapGuides`/`hideSnapGuides`) overlay the snap position during the interaction and are removed on mouseup/touchend. The `#snap-guides` container is lazily created inside `#window-area` with `pointer-events: none`. Layout functions (tile, grid, cascade) reposition all visible windows. Windows reposition/resize proportionally when the browser window or console panel is resized (`scaleWindowsToArea()`). Windows track their own sort/filter state and column autofilters. The `windows` array and `tables` object are the two central data structures. Help windows (Manual, About) are singletons — reopening focuses the existing window instead of creating a duplicate.

### Tabbing and docking

Windows can be combined into dock containers via tabbing and docking. A dock container is a `.dock-container` element positioned in `#window-area` like a `.subwindow`, with its own resize handles, maximize/minimize support, and a recursive binary split tree for layout.

**Data model:** `dockContainers[]` array holds dock containers. Each has a `root` DockNode — either a leaf (tab group) or a split node. Split nodes have `direction` ('horizontal'|'vertical'), `ratio` (0.0–1.0), and two children. Leaf nodes have `tabs` (array of window IDs), `activeTab`, and DOM references (`el`, `tabBarEl`, `contentEl`). Window objects gain `dockId` (null or dock container ID), `dockLeaf` (reference to containing leaf), and `dockable` (false for help/about windows).

**Tabbing:** Hold Shift and drag a window's titlebar onto another window's titlebar area to merge them into a tab group. Tabs are content-width (fit their text) rather than stretching equally. Only 1 level of tabbing (tab groups cannot be nested inside tab groups, but each pane in a dock split can be its own tab group). The active tab's `.win-body` and `.win-statusbar` are visible; inactive tabs' content is hidden via `display: none` with `dataset.winId` for identification. Double-clicking a tab toggles maximize/restore on the dock container. Single-tab leaves inside a split hide their tab bar (class `single-tab`) but the tab shrinks to content width; the tab bar reappears when a second tab is added. Tabs can be reordered by dragging left/right within the same tab bar (no Shift needed) — the dragged tab dims and a blue insertion indicator shows the drop position. The tab bar has an `::after` pseudo-element providing a minimum 40px grab area to the right of tabs so the dock container can always be moved by dragging the tab bar.

**Docking:** Hold Shift and drop a window onto the body area of another window to create a split dock. Without Shift, dragging just moves the window normally. Drop zone detection uses diagonal-based regions: the body area (below titlebar/tab-bar) is divided into 4 triangular zones by its diagonals — the cursor's position relative to the diagonals determines top/right/bottom/left. Dropping on the titlebar or tab bar triggers tabbing (center zone). The overlay shows rectangular half-regions as visual feedback. `getDropZone(x, y, rect, titlebarHeight)` implements the diagonal math using normalized coordinates: `u = bx/bw, v = by/bh` where `bx/by` are relative to the body top-left. Multi-level recursive splitting is supported. The `.dock-splitter` divider is draggable to resize panes (ratio clamped to [0.1, 0.9], double-click resets to 0.5).

**Ghost drag:** Both standalone windows and docked tabs use a ghost element for dock interactions. Shift initiates ghost mode at the 5px movement threshold — once activated, releasing Shift does NOT cancel the mode (the interaction continues until mouseup). The ghost matches the source size (window or leaf pane dimensions) and tracks the cursor with the correct offset from the original mousedown position. For standalone windows, the original window stays in place during ghost drag. Dropping on a valid target executes the dock; dropping on empty space (standalone) is a no-op, or (tab) undocks if the cursor is outside the source dock container. Self-drop (same leaf, center zone) is always a no-op. Without Shift, dragging a tab in a multi-tab leaf enters reorder mode instead of moving the dock — the tab is reordered within its leaf's `tabs` array. Single-tab leaves still move the dock container on plain drag.

**Undocking:** Hold Shift and drag a tab outside its dock container to undock it as a standalone window. The undocked window retains the pane's dimensions and is placed at the drop position with the same mouse offset as the original pickup point.

**Dissolve logic:** When a tab is closed or undocked, `removeTabFromLeaf` handles cleanup. If a leaf becomes empty, `collapseLeaf` removes it from the tree (sibling replaces the parent split). If the entire dock reduces to a single leaf with a single tab, `dissolveDock` converts it back to a standalone window at the dock's position/size.

**Key functions:** `createDockContainer`, `destroyDockContainer`, `renderDockTree`, `buildDockNodeDOM`, `mergeWindowsIntoTabs`, `mergeWindowsAsSplit`, `dockWindowAsTab`, `dockWindowAsSplit`, `undockWindow`, `removeTabFromLeaf`, `collapseLeaf`, `dissolveDock`, `activateTab`, `renderTabBar`, `setupTabDragFromMousedown`, `setupSplitterDrag`, `detectDropTarget`, `findLeafAtPoint`, `getDropZone`, `showDropOverlay`, `hideDropOverlay`, `executeDock`, `toggleMaximizeDock`, `closeDock`, `getLayoutUnits`, `isWindowMinimized`

**Integration:** `focusWindow` brings the dock container to front and activates the tab. `closeWindow`, `minimizeWindow`, `restoreWindow`, `toggleMaximize`, `nudgeWindow` all delegate to the dock container when the window is docked. Layout functions use `getLayoutUnits()` which includes both standalone windows and dock containers. `scaleWindowsToArea` and `getSnapEdges` include dock containers. `updateWindowTitle` and `startInlineRename` update dock tab titles. Help/about/plugin-about windows are created with `dockable: false`. The dock container's `mousedown` handler walks from `e.target` up to the nearest `.dock-leaf` DOM element, matches it to the dock tree, and calls `focusWindow()` with that leaf's `activeTab` — this ensures `activeWinId` stays correct in split layouts so clipboard operations and `getActiveDataWindow()` target the right pane.

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

**IMPORTANT:** `csvsql/static/` contains copies of the root static files (`app.js`, `index.html`, `style.css`, `lib/`, `example/`). These are NOT auto-synced — whenever you modify any root static file, you MUST copy it to `csvsql/static/` as well, or the `csvsql`/`csvsqlw` command will serve stale files.

```bash
# Sync static files to PyPI package (run after ANY change to root static files)
cp app.js index.html style.css csvsql/static/ && cp lib/* csvsql/static/lib/ && cp -r example csvsql/static/

# Build and publish
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
- Cell selection is a rectangle defined by `win.anchorCell` (fixed corner) and the focused cell (lead). Selected cells are cached as a Set of `"rownum:colName"` keys in `win.selectedCells`. Keys use column *names* + `_rownum` so the selection survives sort, filter, column reorder, and rename. `applyCellHighlights()` rebuilds `.row-highlight`/`.col-highlight`/`.cell-selected` classes from this Set after every `renderVisibleRows()` (needed for virtual scrolling). Anchor is reset by user-initiated focusin on a cell; programmatic focus (e.g., drag-end, arrow-key move) sets `win._programmaticFocus` to skip the reset. `selectAllCells(win)` toggles: if all cells are already selected (`isAllSelected(win)`), it calls `selectNoneCells(win)` to clear the selection; otherwise it selects every displayed cell and sets `win._copyWithHeader = true`. `isAllSelected(win)` returns true when `win.selectedCells.size` equals `displayRows * columns`. `selectNoneCells(win)` clears `selectedCells`, `anchorCell`, and `_copyWithHeader`. The Edit menu label toggles between "Select All" and "Select None" via `updateMenuState()`. `selectRows(win, fromDi, toDi)` selects entire row ranges and also sets `_copyWithHeader`. Clicking the `#` corner cell (`th.row-num-header`) triggers select-all/none toggle; clicking a `td.row-num` cell triggers row selection (with drag and Shift+click support). `_copyWithHeader` is cleared by the focusin handler on any normal cell click
- Clipboard: `copySelectedCells(win)` reads raw values from `win.selectedCells`, builds TSV (tab-separated columns, newline-separated rows), and writes to `navigator.clipboard`. When `win._copyWithHeader` is true, the column header row is prepended. `cutSelectedCells(win)` copies then clears selected cells (undoable). `pasteAtAnchor(win)` reads TSV from clipboard and fills right and down from `win.anchorCell`, clipping to table bounds. All three are in the per-table keydown handler (Ctrl/Cmd+C/X/V, select mode only — edit mode passes through to native browser behavior). After paste/cut/undo/redo re-render the table, `refocusAnchorCell(win)` re-focuses the anchor cell's new DOM node so the per-table keydown handler continues to receive events
- Undo/redo stacks (`t._undoStack`, `t._redoStack`) live on the table object (not the window) and are lazily initialized. Each entry is `{ type: 'edit'|'paste'|'cut', changes: [{ rownum, col, oldValue, newValue }] }`. Single-cell edits push from the blur handler; paste and cut push as one entry with all changed cells. `undoTable(tableName, win)` pops from undo, restores old values, pushes to redo; `redoTable` is the reverse. New edits clear the redo stack. The Edit menu shows contextual labels ("Undo Paste", "Redo Cut", etc.) and grays out items when stacks are empty
- Cell editability is a two-state model. Data cells render as `<td class="data-cell" tabindex="0">` — focusable for selection but **not** editable. `enterEditMode(td)` sets `contenteditable="true"` and places the caret at the end; `exitEditMode(td)` removes the attribute. Enter, `i`, F2, Ctrl/Cmd+U, Ctrl/Cmd+click, and touch double-tap are the entry points. The blur handler saves + exits edit mode; Escape either reverts the edit (in edit mode) or clears selection (in select mode). Keep the `data-cell` class as the identity check — don't rely on the presence of `contenteditable` to identify data cells
- Arrow-key behavior on a data cell depends on mode: in select mode, plain arrow collapses the selection to the adjacent single cell (`moveSingleCellSelection`), Shift+arrow extends the rectangle (`extendCellSelection`), and Ctrl/Cmd+←/→ moves the selection's columns as a block (`moveSelectionColumns`). Vim keys h/j/k/l are remapped to ArrowLeft/Down/Up/Right before the arrow check (only in select mode, no ctrl/meta/alt), so they funnel through the same `navKey` path as arrows — Shift+H/J/K/L extends. Ctrl/Cmd+H/J/K/L are claimed by `nudgeWindow` (move the active window 5 px in the vim direction), and Ctrl/Cmd+Shift+L/H by `cycleTableWindow` (cycle prev/next visible data window, skipping minimized and non-table windows) — both deliberately bypass the vim nav remap. Tab/Shift+Tab also cycle windows in select mode; in edit mode Tab still walks cells in the row. In edit mode none of the nav keys fire — arrows and letters pass through to native contenteditable caret movement / typing
- Column reorder: plain drag on a header reorders (5 px threshold); Ctrl/Cmd+click renames; Ctrl/Cmd+←/→ nudges the header-selected column (`win.selectedCol`) when focus is outside a data cell, or the range of columns spanned by the cell selection when a data cell has focus. All paths funnel through `reorderColumn(win, fromIdx, toIdx)`, which splices `tables[name].columns`, re-registers the SQLite table, and calls `rebuildTable(win)`
- Column resize: uses a "measure-then-fix" pattern — on first render, columns auto-size in `table-layout: auto`, then `buildTableHTML` measures each `<th>` offsetWidth, stores in `win.colWidths` (array of px values), and switches to `table-layout: fixed` with a `<colgroup>` controlling widths. Each `<th>` has a `.col-resize-handle` (absolute-positioned div, right edge) whose mousedown calls `startColResize()` — updates `<col>` width and `win.colWidths` on mousemove without rebuilding the table. The row-number column (`th.row-num-header`) also has a resize handle — its width is stored separately in `win.rowNumWidth` (default 50, min 30) and managed by `startRowNumResize()` / `autoFitRowNumColumn()`. `updateTableWidth(win)` keeps the table's inline width in sync with `(win.rowNumWidth || 50) + sum(colWidths)`. Double-click calls `autoFitColumn()` for data columns (measures header + visible rows via off-screen span) or `autoFitRowNumColumn()` for the row-number column (measures largest row number at 11px font). `autoFitAllColumns(win)` fits every data column and the row-number column in a single pass with one measurer element — respects plugin display transforms (formatted vs raw based on `win.disabledTransforms`), caps each column at 75% of the window width (dock leaf pane width if docked), and calls `updateTableWidth` once at the end. Both `autoFitColumn` and `autoFitAllColumns` add `CELL_BORDER` (2px) to data cell measurements to account for the cell's `border: 1px solid` under `box-sizing: border-box`. `reorderColumn()` splices `win.colWidths` in sync. `addColumn()` resets `win.colWidths = null` to re-measure (does not reset `rowNumWidth`). The resize handle sets `th._didDrag = true` on mouseup so the subsequent click doesn't trigger sort (or select-all for the row-num header). The autofilter dropdown includes an `.autofilter-sizing` section (below the Apply/Clear buttons, separated by `border-top`) with "Auto Fit This Column" and "Auto Fit All Columns" buttons
- Column autofilter: each column header has a `☰` button (`.col-filter-btn`) that opens an Excel-style dropdown (`openAutoFilter()`). State: `win.columnFilters` is `{ colName: Set<string> }` — absent key means no filter. `_activeAutoFilter` (module-level) tracks the single open dropdown `{ win, col, el, scrollHandler }`. The dropdown lists unique values from ALL rows (not filtered rows), with a search box, Select All, and Apply/Clear buttons. Column filters are applied in `buildTableHTML()` after the SQL WHERE filter and before sorting — they AND together. `closeAutoFilter()` is called by `buildTableHTML`, resize handles, window resize, outside click, and Escape. `renameColumn()` migrates filter entries. Headers use a `.th-inner` flex wrapper (not `display:flex` on `th` itself, which breaks table layout). Filtered headers get `.col-filtered` (green left border) and the icon gets `.active` (always visible). Virtual scrolling kicks in for 200+ unique values. When any filter is active (column autofilters or WHERE text), the status bar shows a "Clear Filters" link (`.status-clear-filters`) that resets `win.columnFilters`, `win.filterText`, and the filter input in one click
- Plugin system: plugins are JSON config files loaded at runtime that control per-column display formatting via the CSVSQL expression language (a sandboxed evaluator with no JS execution) and cross-table linking. Two formats: legacy (`table`/`columns` at top level) and multi-table (`tables` array, each with `table`/`columns`). Optional metadata fields: `version`, `author`, `created`, `description` — shown in the plugin About dialog. State: `plugins[]` (loaded configs with compiled ASTs), `_columnTransformCache` (tableName → colName → displayAst, rebuilt on plugin load/unload, table open/close/rename, column rename), `_linkCache` (compiled link rules). Key functions: `exprTokenize/exprParse/exprEval` (expression language), `loadPluginFile/loadPluginFromFile/unloadPlugin` (lifecycle), `rebuildTransformCache/rebuildTransformCacheForTable/rebuildLinkCache` (cache), `getDisplayValue/hasDisplayTransform` (render-time lookup), `applyLinkFilters/clearLinkFilters` (cross-table linking), `showPluginAbout` (About dialog modal), `showToast` (non-blocking toast notifications). `renderVisibleRows()` calls `getDisplayValue()` instead of raw `row[col]`. `enterEditMode()` restores the raw value when a display transform is active. Plugins persist in localStorage under `csvsql_plugins`. Bundled examples in `example/`. Drag-and-drop of `.json` files routes to `loadPluginFile()` instead of `openFileByType()`. Plugin menu: each loaded plugin shows as a `.plugin-entry` div with a `.plugin-unload` (✕) span for quick unload (the menu stays open so multiple plugins can be unloaded without reopening) and a `.plugin-name` span that opens the About dialog (`showPluginAbout`). The About dialog is a subwindow (via `showPluginAbout` → `createSubwindow`) with plugin metadata, column rules, link rules, and an Unload button. It is a singleton tracked by `_activePluginAboutWin` — reopening creates a new window and closes the old one. Toast notifications (`showToast(message, type)`) replace `alert()` for plugin load/unload feedback — `.toast-success` (green) and `.toast-error` (red), auto-dismiss after 3s with slide animation. Status bar toggle chips: the status bar center area shows toggle chips (Sort, Filter, Link, Format) for features that are active on the window. Each chip is clickable to suspend/resume that feature. Per-window state: `win.disableSort`, `win.disableFilter`, `win.disableLink` (booleans), `win.disabledTransforms` (Set of column names). Column headers get `.col-transformed` class when a plugin transform is active, with `.feature-disabled` when suspended. The Format chip toggles all transform columns in/out of `disabledTransforms`. Keyboard shortcuts: Ctrl/Cmd+Shift+1/2/3/4 toggle Sort/Filter/Link/Format. CSS: `.status-chip` with `.status-chip-sort/filter/link/format` for colors, `.off` class when suspended. Column left-borders indicate feature state (purple=sorted, green=filtered, blue=linked, pink=formatted), dashed+dimmed when that feature is suspended. All 4 transform call sites (renderVisibleRows, blur handler, Escape revert, enterEditMode) guard with `!win.disabledTransforms.has(col)`. The autofilter dropdown (`openAutoFilter`) shows display-formatted values when a transform is active and enabled for the column; the underlying filter Set stores raw values
- Cross-table linking: plugins can define a `links` array where each entry has `source: { table, column }` and `target: { table, column }` (all regex patterns, anchored like display rules). When rows are selected in a source table, `applyLinkFilters()` (called from `applyCellHighlights()`) collects distinct values from matched source columns and sets `win.linkFilters` on target windows — a separate filter object from `win.columnFilters` so manual filters and link filters are independent. Link filters are applied in `buildTableHTML()` after column autofilters. The source table is always excluded from its own link targets (even with `.*` patterns). Clearing selection calls `clearLinkFilters()` to remove link filters. Unloading a plugin clears all link filters. Visual indicators: `.col-linked` class on link-filtered column headers (blue left border), Link chip (`.status-chip-link`) in the status bar. Link filters propagate transitively via BFS — selecting rows in table A filters table B, and the matching rows in B are used to filter table C, and so on. A `visited` set (keyed by table name) prevents cycles, and a depth cap of 10 stops runaway chains. A re-entrancy guard (`_applyingLinkFilters`) prevents recursive calls triggered by `rebuildTable` callbacks. Tables that are not link sources don't interfere with existing link filters when clicked
- Touch gestures are layered on top of the mouse handlers and all use touch events only (no pointer-event fallback). State lives at IIFE scope: `_touchHeaderDrag` (active header drag), `_touchCellDrag` (active cell second-tap interaction — either double-tap edit or 1.5-tap multi-select), `_touchWinDrag` (active titlebar window drag), and the tap-one trackers `_lastHeaderTap` / `_lastCellTap` / `_lastTitleTap` (each keyed by a stable identifier — e.g. table + column name, window id + rownum + col — because `rebuildTable` recreates DOM nodes between taps). The cell gesture is a single branching interaction: first tap just records the cell identity and time; on second tap within `DOUBLE_TAP_WINDOW_MS` (500 ms) we enter `_touchCellDrag` with `dragging: false`, seed `win.anchorCell` from the first-tap cell, then disambiguate at the gesture end — if touchmove exceeded 5 px, `dragging` flipped to true and we finalize a drag-select rectangle (`rebuildSelectionRect` → `focusCellAt` + preventDefault touchend so the synthesized click doesn't collapse the selection); otherwise it's a double-tap and we flip `contenteditable="true"` on the tapped cell synchronously and do NOT call preventDefault — the touchend-synthesized click then lands on the now-editable cell and iOS opens the virtual keyboard natively (programmatic `focus()` inside a touch handler does not). Data cells set `user-select: none` + `-webkit-touch-callout: none` + `touch-action: manipulation` so iOS/Android don't hijack a still touch for text selection/callout and don't trigger double-tap zoom; edit mode overrides back to `user-select: text`. The header drag sets `th._didDrag = true` so the trailing click doesn't sort
