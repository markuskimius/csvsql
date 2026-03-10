# CSVSQL

A browser-based database that uses CSV files as storage. Open CSV files as tables, run SQL queries against them, and edit data inline — all in a multi-window interface with no server or build step required.

## Features

- **CSV as database tables** — Each CSV file is a table. Open multiple files and query across them.
- **SQL queries** — Write and execute SQL from the built-in console (powered by [AlaSQL](https://github.com/AlaSQL/alasql)). Each query result opens in its own window.
- **Inline editing** — Click any cell to edit. Tab/Enter to navigate, Escape to cancel.
- **Sort and filter** — Click column headers to sort. Use the filter bar to search across all columns.
- **Multi-window workspace** — Every table and query result lives in its own draggable, resizable subwindow. Arrange them with Tile, Grid, or Cascade layouts.
- **Row management** — Add/delete rows, insert rows at any position (right-click row numbers), add new columns.
- **Save** — Export any table or query result back to CSV.

## Getting Started

No dependencies to install. Just serve the files or open directly:

```sh
# Option 1: any static file server
python3 -m http.server 8000
# then open http://localhost:8000

# Option 2: open directly
open index.html
```

Use **File > Open CSV** to load a CSV file (a `sample.csv` is included), or **File > New Table** to create one from scratch.

## SQL Console

The console sits at the bottom of the window. Type a query and press **Ctrl+Enter** (or click Run) to execute.

```sql
-- Query a loaded table by its filename (minus extension)
SELECT * FROM sample WHERE name LIKE 'A%'

-- Aggregate queries
SELECT department, AVG(salary) as avg_salary FROM employees GROUP BY department

-- Join across tables
SELECT a.name, b.value FROM table1 a JOIN table2 b ON a.id = b.id
```

Table names are derived from the CSV filename (e.g., `employees.csv` becomes `employees`).

## Keyboard Shortcuts

| Key | Action |
|---|---|
| Ctrl+Enter | Execute SQL query |
| Tab / Shift+Tab | Navigate between cells in a row |
| Enter | Move to next row (same column) |
| Escape | Cancel cell edit |
