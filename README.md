# DuckSheet

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![.NET 8](https://img.shields.io/badge/.NET-8.0-purple.svg)](https://dotnet.microsoft.com/en-us/download/dotnet/8)
[![ExcelDNA](https://img.shields.io/badge/ExcelDNA-1.9.0-green.svg)](https://excel-dna.net/)
[![DuckDB](https://img.shields.io/badge/DuckDB-1.2.1-yellow.svg)](https://duckdb.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20x64-lightgrey.svg)]()

**Push Excel ranges into DuckDB. Query them with SQL. All from worksheet formulas.**

DuckSheet is an Excel XLL add-in that exposes five worksheet functions. No ribbon. No task pane. No configuration files. Your formulas are the config.

---

## What It Does

```
Excel range  ──►  DUCK.SEND  ──►  DuckDB table
                  DUCK.QUERY ──►  results back to any cell or sheet
                  DUCK.EXECUTE ──►  DDL, DML, views, transforms
```

Write data from Excel into a local `.duckdb` file, then query it with full SQL — joins, aggregations, window functions, anything DuckDB supports. External tools like DBeaver can connect to the same file read-only while DuckSheet is running.

---

## Screenshots

> 📸 **Control panel** — all DUCK formulas in one sheet, results written elsewhere

![Control Panel](docs/control-panel.png)

> 📸 **DUCK.SEND** — pushing a range into DuckDB with type inference

![Send](docs/duck-send.png)

> 📸 **DUCK.QUERY** — SQL results written directly to the Results sheet

![Query](docs/duck-query.gif)

---

## Requirements

| Requirement | Version |
|---|---|
| Windows | 10 / 11 (x64) |
| Excel | 2016 or later (64-bit) |
| .NET Runtime | [.NET 8](https://dotnet.microsoft.com/en-us/download/dotnet/8) |

> **.NET 8 must be installed** on the target machine. It is not bundled with the add-in.

---

## Installation

1. Download the latest release from the [releases](release/) folder:
   - `DuckSheet-AddIn64-packed.xll`
   - `duckdb.dll`

2. Place **both files in the same folder**.

3. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Browse** → select the `.xll` file.

4. Done. The five `DUCK.*` functions are now available in any workbook.

---

## The Five Functions

### `DUCK.CONNECT` — register a database
```excel
=DUCK.CONNECT("C:\data\mywork.duckdb", "db1")
```
Creates the `.duckdb` file if it doesn't exist, or connects to an existing one.
Registers it under a short name used by all other functions.
Multiple databases can be registered simultaneously.

**Returns:** `Registered: db1 → mywork.duckdb  3/17/2026 19:38`

---

### `DUCK.SEND` — push a range to DuckDB
```excel
=DUCK.SEND(A1:F500, "sales", "db1")
=DUCK.SEND(Data!A1:F500, "sales", "db1")
```
- First row = column headers
- Drops and recreates the table on every call (overwrite semantics)
- Detects Excel date-formatted columns automatically
- Empty cells, errors, blanks → NULL

**Returns:** Multi-line status with row/col counts and inferred types per column

---

### `DUCK.QUERY` — run a SELECT, write results to a cell
```excel
=DUCK.QUERY("SELECT * FROM sales", "Results!$A$1", "db1")
=DUCK.QUERY("SELECT * FROM sales WHERE region = 'West'", "$E$1", "db1")
=DUCK.QUERY("C:\queries\monthly_report.sql", "Results!$A$1", "db1")
```
- Results (with headers) are written starting at the target cell
- Target is always a **plain string address** — supports same-sheet and cross-sheet
- Pass a `.sql` or `.txt` file path instead of inline SQL
- Never use `CELL()` as the target — it is volatile and will loop

**Returns:** `Query OK: 99 rows → Results!$A$1  3/17/2026 19:38`

---

### `DUCK.EXECUTE` — DDL and DML
```excel
=DUCK.EXECUTE("CREATE OR REPLACE VIEW summary AS SELECT ...", "db1")
=DUCK.EXECUTE("DELETE FROM staging WHERE processed = true", "db1")
=DUCK.EXECUTE("C:\queries\build_views.sql", "db1")
```
- Use for anything that doesn't return rows: CREATE, DROP, INSERT, ALTER, PRAGMA
- Pass a `.sql` or `.txt` file path instead of inline SQL

**Returns:** `Created view: summary  3/17/2026 19:38` / `Deleted: 42 rows  3/17/2026 19:38`

---

### `DUCK.INSTRUCTIONS` — built-in reference
```excel
=DUCK.INSTRUCTIONS()
```
Returns full usage documentation as a string. View it by expanding the formula bar or looking at the cell tooltip.

---

## Control Panel Pattern

Place all DUCK formulas in a dedicated sheet. Excel evaluates top-to-bottom so CONNECT must be above SEND and QUERY.

```
[DuckDB sheet]

B2:  =DUCK.CONNECT("C:\data\mywork.duckdb", "db1")

B4:  =DUCK.SEND(Data!A1:F500, "sales", "db1")
B5:  =DUCK.SEND(Data!H1:K200, "products", "db1")

B7:  =DUCK.EXECUTE("C:\queries\build_views.sql", "db1")

B9:  =DUCK.QUERY("SELECT * FROM summary", "Results!$A$1", "db1")
B10: =DUCK.QUERY("C:\queries\rep_report.sql", "Results!$A$20", "db1")
```

**Refresh a single formula:** click the cell → `F2` → `Enter`
**Refresh everything:** `Ctrl+Alt+F9`

---

## Type Inference

`DUCK.SEND` scans each column and picks the narrowest type that fits all values:

| DuckDB Type | Inferred When |
|---|---|
| `BOOLEAN` | All non-null values are true/false |
| `TIMESTAMP` | All non-null values parse as dates, or cell has a date number format |
| `BIGINT` | All non-null values are whole numbers |
| `DOUBLE` | All non-null values are numeric |
| `VARCHAR` | Fallback |

Mixed columns (e.g. mostly numbers with some "N/A") → `DOUBLE`, strings become NULL.
Excel date-formatted cells (`MM/DD/YYYY` etc.) are detected automatically and stored as `TIMESTAMP`.

---

## File-Based SQL

Both `DUCK.QUERY` and `DUCK.EXECUTE` accept a file path instead of inline SQL.
If the first argument ends in `.sql` or `.txt`, the file is read and used as the query.

```excel
=DUCK.QUERY("C:\queries\monthly_report.sql", "Results!$A$1", "db1")
=DUCK.EXECUTE("C:\queries\build_views.sql", "db1")
```

This keeps complex SQL out of formula cells and lets you edit queries in a proper SQL editor.

---

## DBeaver Integration

```
Driver:    DuckDB
File:      same path used in DUCK.CONNECT
Property:  access_mode = read_only
```

DuckSheet opens and closes its connection on every function call, so DBeaver can connect freely between operations — no lock contention.

---

## Describe a Table

```excel
=DUCK.QUERY("DESCRIBE sales", "Results!$H$1", "db1")
=DUCK.QUERY("SHOW TABLES", "Results!$H$1", "db1")
```

---

## Building from Source

```bash
git clone https://github.com/zz415/ducksheet.git
cd ducksheet
dotnet build -c Release
```

Output is in `bin\Release\net8.0-windows\publish\`:
- `DuckSheet-AddIn64-packed.xll`
- `duckdb.dll`

Both files must stay in the same folder.

---

## License

[MIT](LICENSE) © 2026 zz415
