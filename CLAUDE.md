# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

CSVSQL is a browser-based CSV database application. It treats CSV files as database tables with SQL query support, editable cells, and a multi-window interface. No build step or server required — open `index.html` directly in a browser.

## Architecture

Single-page app with three files:

- **index.html** — Shell: menubar, workspace area, SQL console panel, hidden file input
- **style.css** — Dark theme styling, window management visuals, table layout
- **app.js** — All application logic in a single IIFE (`app` module), exposing methods on the global `app` object

### Key dependencies (loaded via CDN)

- **Papa Parse** — CSV parsing/unparsing
- **AlaSQL** — In-browser SQL engine; each opened CSV is registered as an AlaSQL table

### Data flow

1. CSV opened via File menu → Papa Parse parses → stored in `tables[name]` object (columns, rows, filename, modified flag)
2. Table registered in AlaSQL with `CREATE TABLE` + data injected into `alasql.tables[name].data`
3. Edits to cells update `tables[name].rows` and sync back to AlaSQL via `syncToAlaSQL()`
4. SQL queries run against AlaSQL; results become new entries in `tables` and open in new subwindows
5. Save serializes `tables[name]` back to CSV via Papa.unparse and triggers browser download

### Window management

Custom subwindow system — each table/query result gets a draggable, resizable, minimizable window inside `#window-area`. Layout functions (tile, grid, cascade) reposition all visible windows. Windows track their own sort/filter state. The `windows` array and `tables` object are the two central data structures.

### Row identity

Each row's primary key is its `_rownum` property (1-based index). Renumbered on insert/delete. Not a real column — excluded from CSV export.

## Development

No build tools. To develop:

```
# Serve locally (any static server works)
python3 -m http.server 8000
# Then open http://localhost:8000
```

Or just open `index.html` directly — the only external resources are CDN scripts.

## Conventions

- All state lives in the `app` IIFE's closure (`windows`, `tables`, `nextWinId`, etc.)
- Public methods are returned from the IIFE and called from HTML onclick handlers or internally
- Table names are sanitized to `[a-zA-Z0-9_]` for AlaSQL compatibility
- AlaSQL table names use bracket-quoting (`[tableName]`) to handle edge cases
