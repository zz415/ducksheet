# DuckSheet

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![.NET 8](https://img.shields.io/badge/.NET-8.0-purple.svg)](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
[![ExcelDNA](https://img.shields.io/badge/ExcelDNA-1.9.0-green.svg)](https://excel-dna.net/)
[![DuckDB](https://img.shields.io/badge/DuckDB-1.2.1-yellow.svg)](https://duckdb.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20x64-lightgrey.svg)]()

**Push Excel ranges into DuckDB. Query them with SQL. All from worksheet formulas.**

DuckSheet is an Excel XLL add-in that exposes five worksheet functions. No ribbon. No task pane. No configuration files. No VBA. Your formulas are the config — place them in a sheet, hit `Ctrl+Alt+F9`, and your data is in DuckDB.

---

## What It Does

```
Excel range  ──►  DUCK.SEND    ──►  DuckDB table (typed, nulls handled)
SQL string   ──►  DUCK.QUERY   ──►  results written to any cell or sheet
SQL string   ──►  DUCK.EXECUTE ──►  DDL, DML, views, transforms
.sql file    ──►  DUCK.QUERY / DUCK.EXECUTE  ──►  file-based SQL workflows
```

Write data from any Excel range into a local `.duckdb` file, then query it with full SQL — joins, aggregations, window functions, CTEs, anything DuckDB supports. External tools like DBeaver can connect to the same file read-only at any time while DuckSheet is running.

---

## Screenshots

> 📸 **Control panel** — all DUCK formulas in one sheet, results written elsewhere

![Control Panel](docs/control-panel.png)

> 📸 **DUCK.SEND** — pushing a range into DuckDB with automatic type inference

![Send](docs/duck-send.png)

> 📸 **DUCK.QUERY** — SQL results written directly to the Results sheet

![Query](docs/duck-query.gif)

---

## Requirements

| Requirement | Details |
|---|---|
| Windows | 10 or 11 (x64) |
| Excel | 2016 or later, 64-bit |
| [.NET 8 Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) | Must be installed separately — not bundled |

> **.NET 8 must be installed** on the target machine before loading the add-in.
> Download it from: https://dotnet.microsoft.com/en-us/download/dotnet/8.0

---

## Installation

1. Download both files from the [`release/`](release/) folder:
   - `DuckSheet-AddIn64-packed.xll`
   - `duckdb.dll`

2. Place **both files in the same folder** — they must stay together.

3. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Browse** → select the `.xll` file → OK.

4. Done. The five `DUCK.*` functions are now available in every workbook.

> On first open, Windows may show a security warning. Click **Enable** or **Open** to allow the add-in to load.

---

## The Five Functions

### `DUCK.CONNECT` — register a database file

```excel
=DUCK.CONNECT("C:\data\mywork.duckdb", "db1")
```

Registers a `.duckdb` file under a short name. All other functions use that name to know which database to talk to.

- Creates the `.duckdb` file on disk if it does not already exist
- If the file exists, connects to it as-is — existing tables and views are preserved
- The connection is tested immediately; if the path is invalid you get an error in the cell
- Multiple databases can be registered at the same time using different names:

```excel
B2: =DUCK.CONNECT("C:\data\sales.duckdb",   "sales")
B3: =DUCK.CONNECT("C:\data\hr.duckdb",      "hr")
B4: =DUCK.CONNECT("C:\data\finance.duckdb", "fin")
```

Each function call that follows will reference whichever name it needs. You can mix databases freely in the same workbook.

**Returns:** `Registered: db1 → mywork.duckdb  3/17/2026 19:38`

---

### `DUCK.SEND` — push an Excel range into DuckDB

```excel
=DUCK.SEND(A1:F500, "sales", "db1")
=DUCK.SEND(Data!A1:F500, "sales", "db1")
```

Reads an Excel range and writes it into a DuckDB table. The first row must be column headers. Every subsequent call drops the table and recreates it — overwrite semantics, no appending.

**Type inference** runs per column, scanning all non-null values and picking the narrowest type that fits:

| DuckDB Type | Inferred When |
|---|---|
| `BOOLEAN` | All non-null values are `TRUE` / `FALSE` |
| `TIMESTAMP` | All non-null values parse as dates, or the cell has a date number format |
| `BIGINT` | All non-null values are whole numbers |
| `DOUBLE` | All non-null values are numeric (including decimals) |
| `VARCHAR` | Fallback — anything that doesn't fit above |

**Date detection:** Excel date cells are stored internally as floating-point serial numbers — there is no way to distinguish them from regular numbers by value alone. DuckSheet reads each column's number format directly from Excel. If a column's format contains a year token (`y`), it is treated as `TIMESTAMP` and converted via `DateTime.FromOADate()`. This means date columns with mixed nulls and errors are handled correctly.

**Null handling:** Empty cells, Excel errors (`#N/A`, `#REF!`, etc.), and blank strings all become SQL `NULL`. No value is ever invented to fill a row. NULL counts per column are surfaced in the return string.

**Returns:** Multi-line status:
```
Loaded: sales — 499 rows, 6 cols
  order_id: VARCHAR
  order_date: TIMESTAMP
  quantity: BIGINT (3 nulls)
  unit_price: DOUBLE
  is_renewal: BOOLEAN
  region: VARCHAR  3/17/2026 19:38
```

---

### `DUCK.QUERY` — run a SELECT and write results to a cell

```excel
=DUCK.QUERY("SELECT * FROM sales", "Results!$A$1", "db1")
=DUCK.QUERY("SELECT region, SUM(revenue) FROM sales GROUP BY region", "$E$1", "db1")
=DUCK.QUERY("C:\queries\monthly_report.sql", "Results!$A$1", "db1")
```

Executes a SQL SELECT statement and writes the results — with column headers in the first row — starting at the target cell address. The target can be on any sheet in the workbook.

**Target address rules:**
- Always a plain string — type it directly or reference a named range
- Supports same-sheet and cross-sheet addresses:

```excel
"E1"           -- same sheet, relative
"$E$1"         -- same sheet, absolute
"Results!A1"   -- another sheet named Results
"Results!$A$1" -- another sheet, absolute (recommended)
```

> ⚠️ **Never use `CELL("address", E1)` as the target.** `CELL()` is a volatile function. Every time DuckSheet writes results to the sheet, Excel recalculates all volatile functions — including `CELL()` — which re-runs the query, which writes again, causing an infinite loop.

**File-based SQL:** If the first argument ends in `.sql` or `.txt`, DuckSheet reads that file and uses its contents as the query. This keeps complex SQL out of formula cells and lets you edit queries in a proper SQL editor.

```excel
=DUCK.QUERY("C:\queries\monthly_report.sql", "Results!$A$1", "db1")
=DUCK.QUERY("C:\queries\my_query.txt", "$E$1", "db1")
```

**Returns:** `Query OK: 99 rows → Results!$A$1  3/17/2026 19:38`

---

### `DUCK.EXECUTE` — run DDL or DML

```excel
=DUCK.EXECUTE("CREATE OR REPLACE VIEW summary AS SELECT region, SUM(revenue) FROM sales GROUP BY region", "db1")
=DUCK.EXECUTE("DELETE FROM staging WHERE processed = true", "db1")
=DUCK.EXECUTE("C:\queries\build_views.sql", "db1")
```

Executes a SQL statement that performs an action but returns no rows. Use this for anything structural or transformational — building views, dropping tables, inserting records, running PRAGMA commands.

Use `DUCK.QUERY` for anything that returns rows. Use `DUCK.EXECUTE` for everything else.

**File-based SQL:** Same as `DUCK.QUERY` — if the first argument ends in `.sql` or `.txt`, the file is read and used as the SQL. This is especially useful for multi-statement setup scripts:

```excel
=DUCK.EXECUTE("C:\queries\build_views.sql", "db1")
```

Where `build_views.sql` might contain multiple `CREATE OR REPLACE VIEW` statements separated by semicolons.

**Returns:** A description of what happened, parsed from the SQL verb:

```
Created table: orders       3/17/2026 19:38
Created or replaced view: summary  3/17/2026 19:38
Dropped table: staging      3/17/2026 19:38
Deleted: 42 rows            3/17/2026 19:38
Updated: 7 rows             3/17/2026 19:38
OK                          3/17/2026 19:38   ← fallback for unrecognised statements
```

---

### `DUCK.INSTRUCTIONS` — built-in reference

```excel
=DUCK.INSTRUCTIONS()
```

Returns the full usage reference as a string. Expand the formula bar or widen the cell to read it. No parameters needed.

---

## Control Panel Pattern

The recommended layout is a dedicated sheet (e.g. `DuckDB`) containing all DUCK formulas, with results written out to separate sheets. Excel evaluates formulas top-to-bottom, so `DUCK.CONNECT` must appear above any `DUCK.SEND` or `DUCK.QUERY` that depends on it.

```
[DuckDB sheet]

B2:  =DUCK.CONNECT("C:\data\mywork.duckdb", "db1")

B4:  =DUCK.SEND(Data!A1:F500,   "sales",    "db1")
B5:  =DUCK.SEND(Data!H1:K200,   "products", "db1")
B6:  =DUCK.SEND(People!A1:O501, "staff",    "db1")

B8:  =DUCK.EXECUTE("C:\queries\build_views.sql", "db1")

B10: =DUCK.QUERY("SELECT * FROM summary",         "Results!$A$1",  "db1")
B11: =DUCK.QUERY("C:\queries\rep_report.sql",     "Results!$A$20", "db1")
B12: =DUCK.QUERY("SELECT * FROM staff_summary",   "Results!$A$40", "db1")
```

**Refresh a single formula:** click the cell → `F2` → `Enter`
**Refresh everything:** `Ctrl+Alt+F9` — recalculates all formulas in the workbook top to bottom

---

## Describe a Table

DuckDB's native `DESCRIBE` and `SHOW` work directly through `DUCK.QUERY`:

```excel
=DUCK.QUERY("DESCRIBE sales",  "Results!$H$1", "db1")
=DUCK.QUERY("SHOW TABLES",     "Results!$H$1", "db1")
=DUCK.QUERY("SHOW ALL TABLES", "Results!$H$1", "db1")
```

---

## DBeaver Integration

DuckSheet opens and closes its database connection on every function call. This means external tools can connect to the same `.duckdb` file at any time — no lock contention, no need to close Excel first.

```
Driver:   DuckDB
File:     same path used in DUCK.CONNECT
Property: access_mode = read_only
```

Set `access_mode = read_only` to prevent DBeaver from accidentally writing to the file while DuckSheet is using it.

---

## Building from Source

```bash
git clone https://github.com/zz415/ducksheet.git
cd ducksheet
dotnet build -c Release
```

Output lands in `bin\Release\net8.0-windows\publish\`:
- `DuckSheet-AddIn64-packed.xll` — the add-in (all managed assemblies packed inside)
- `duckdb.dll` — native DuckDB binary, must stay in the same folder as the XLL

Both files are required. Do not separate them.

---

## License

[MIT](LICENSE) © 2026 zz415
