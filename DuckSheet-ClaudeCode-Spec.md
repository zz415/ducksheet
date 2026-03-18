# DuckSheet — Project Specification (Current State)

## What This Is

DuckSheet is an ExcelDNA XLL add-in that exposes worksheet functions allowing users to push Excel cell ranges into local DuckDB database files and query that data with SQL. Multiple databases can be registered and used simultaneously.

**SQL is a first-class citizen. Excel is the data source. DuckDB is the engine.**

---

## Tech Stack

| Layer | Choice |
|---|---|
| Add-in framework | ExcelDNA 1.9.0 (XLL) |
| Language | C#, net8.0-windows, x64 |
| Database engine | DuckDB.NET.Data.Full 1.2.1 (in-process) |
| Database file | File-backed `.duckdb` |
| Distribution | `DuckSheet-AddIn64-packed.xll` + `duckdb.dll` side-by-side |
| Runtime | .NET 8 must be installed on target machine |

---

## Architecture

```
Excel Worksheet
     │
     │  DUCK.CONNECT / DUCK.SEND / DUCK.QUERY / DUCK.EXECUTE / DUCK.INSTRUCTIONS
     ▼
ExcelDNA XLL Add-in (C#)
     │
     │  DuckDB.NET (in-process, open/close per call)
     ▼
.duckdb file on disk  ◄──────── DBeaver / external SQL tool (read_only connection)
```

Each function opens a fresh connection, runs, and closes. DBeaver can connect at any time without lock contention.

---

## The Five Functions

### 1. `DUCK.CONNECT`
```excel
=DUCK.CONNECT("C:\data\mywork.duckdb", "db1")
```
- Creates the `.duckdb` file if it doesn't exist, connects to existing if it does
- Registers the file under a name used by all other functions
- Multiple databases can be registered simultaneously
- **Returns:** `Registered: db1 → mywork.duckdb  3/17/2026 19:38`

### 2. `DUCK.SEND`
```excel
=DUCK.SEND(A1:F500, "sales_data", "db1")
```
- First row = column headers
- Drops and recreates the table on every call (overwrite semantics)
- Type inference per column: BOOLEAN → TIMESTAMP → BIGINT → DOUBLE → VARCHAR
- Excel-formatted date columns (cells with a date number format) are automatically detected and stored as TIMESTAMP — no manual intervention needed
- Empty cells, errors, blanks → NULL (never poisons the row)
- **Returns:** Multi-line status with row/col count and per-column types + null counts, ending with timestamp

### 3. `DUCK.QUERY`
```excel
=DUCK.QUERY("SELECT * FROM sales_data", "E1", "db1")
=DUCK.QUERY("SELECT * FROM sales_data", "$E$1", "db1")
=DUCK.QUERY("SELECT * FROM sales_data", "Results!A1", "db1")
=DUCK.QUERY("SELECT * FROM sales_data", "Results!$A$1", "db1")
=DUCK.QUERY("C:\queries\monthly_report.sql", "Results!$A$1", "db1")
=DUCK.QUERY("C:\queries\my_query.txt", "E1", "db1")
```
- Executes a SELECT and writes results (with headers) to the target address
- Target must be a **plain string address** — do NOT use `CELL()`, it is volatile and causes an infinite recalculation loop
- Supports current sheet (`"E1"`, `"$E$1"`) and cross-sheet (`"Results!A1"`, `"Results!$A$1"`) references
- If the `sql` parameter ends in `.sql` or `.txt`, the file is read and its contents used as the query
- **Returns:** `Query OK: N rows → E1  3/17/2026 19:38`

### 4. `DUCK.EXECUTE`
```excel
=DUCK.EXECUTE("CREATE VIEW summary AS SELECT ...", "db1")
=DUCK.EXECUTE("C:\queries\build_views.sql", "db1")
=DUCK.EXECUTE("C:\queries\setup.txt", "db1")
```
- For SQL that performs an action but returns no rows
- Use for: CREATE, DROP, INSERT, ALTER, PRAGMA
- If the `sql` parameter ends in `.sql` or `.txt`, the file is read and its contents used as the SQL
- **Returns:** Action description + timestamp, e.g. `Created table: orders  3/17/2026 19:38` or `Deleted: 42 rows  3/17/2026 19:38`. Falls back to `OK  timestamp` for unrecognised statements.

### 5. `DUCK.INSTRUCTIONS`
```excel
=DUCK.INSTRUCTIONS()
```
- Returns full usage reference as a multi-line string
- No parameters needed
- **Returns:** Formatted instructions string

---

## Describe a Table
DUCK.DESCRIBE was removed. Use DUCK.QUERY with DuckDB's native DESCRIBE:
```excel
=DUCK.QUERY("DESCRIBE sales_data", "H1", "db1")
```

---

## Control Panel Layout

Place all DUCK formulas on a dedicated sheet. Excel evaluates top-to-bottom so CONNECT must be above SEND/QUERY.

```
B2:  =DUCK.CONNECT("C:\data\mywork.duckdb", "db1")
B4:  =DUCK.SEND(Data!A1:F500, "sales", "db1")
B5:  =DUCK.SEND(Data!H1:K200, "products", "db1")
B7:  =DUCK.EXECUTE("CREATE OR REPLACE VIEW summary AS SELECT ...", "db1")
B9:  =DUCK.QUERY("SELECT * FROM summary", "Results!$A$1", "db1")
```

---

## Refresh Behavior

- **Single formula:** Click cell → F2 → Enter
- **Full refresh:** Ctrl+Alt+F9 — recalculates all DUCK formulas top to bottom

---

## DBeaver Integration

```
Driver:    DuckDB
File:      same path used in DUCK.CONNECT
Property:  access_mode = read_only
```

DuckSheet opens and closes connections per call so DBeaver can connect freely between operations.

---

## Type Inference

Pass 1 — scan all non-null values per column.
Pass 2 — narrowest type that all values conform to:

```
BOOLEAN → TIMESTAMP → BIGINT → DOUBLE → VARCHAR
```

- **Excel-formatted date columns:** Before value scanning, `DUCK.SEND` reads the number format of the first non-null double in each column via `xlfGetCell`. If the format contains `'y'` (year token), the column is forced to TIMESTAMP and doubles are converted with `DateTime.FromOADate()`. This correctly handles date columns with nulls and errors — only the first non-null cell is checked.
- TIMESTAMP also inferred from string columns where all non-null values parse as dates via `DateTime.TryParse`
- Integer-looking doubles (e.g. year values `2022`, `2023`) stay BIGINT as long as their cell format is `General` or numeric — no false positive date detection
- Mixed columns (e.g. mostly numeric with some "N/A") → DOUBLE, strings become NULL
- NULL coercions counted and surfaced in DUCK.SEND status

---

## Known Technical Notes

- `duckdb.dll` must be in the same folder as the XLL
- The `.dna` file must be auto-generated by the build — do not manually maintain it
- `EnableDynamicLoading=true` is required in the csproj for .NET 8 framework-dependent deployment
- Self-contained deployment is not supported by ExcelDNA — crashes Excel
- `IsMacroType=true` + `QueueAsMacro` causes infinite recalc loop — DUCK.QUERY uses string address to avoid this; DUCK.SEND uses `IsMacroType=true` safely because it never writes to Excel cells
- `CELL()` is volatile — never use `CELL("address",E1)` as the target parameter for `DUCK.QUERY`; it causes an infinite recalculation loop. Always use a plain string like `"E1"` or `"Results!$A$1"`
- `xlSet` (XLL C API) cannot write to a sheet other than the active one — crashes Excel. `DUCK.QUERY` uses `ExcelDnaUtil.Application` (COM) via `QueueAsMacro` to write results with `Range.Value`, which correctly handles cross-sheet writes and all data types
- `DUCK.SEND` requires `AllowReference=true` on the range parameter to access cell metadata (`xlfGetCell`) for date format detection; falls back gracefully if called with a literal array

---

## Next Steps Under Discussion

1. **net48 target** — .NET Framework 4.8 is built into Windows 10/11, eliminates runtime install requirement entirely, zero deployment friction for mass distribution
2. **JS/Office Add-in** — required path for Microsoft AppSource marketplace listing

---

## Project Structure

```
DuckSheet/
  DuckSheet.csproj          — net8.0-windows, x64, ExcelDNA 1.9.0, DuckDB.NET 1.2.1
  AddIn.cs                  — IExcelAddIn lifecycle, AddDllDirectory for native duckdb.dll
  ConnectionManager.cs      — ConcurrentDictionary<name, path>, open/close per call
  TypeInference.cs          — Two-pass column type scanner
  NullChecker.cs            — Cell value → null classification
  RangeWriter.cs            — QueueAsMacro + ExcelReference.SetValue output helper
  Functions/
    ConnectFunction.cs      — DUCK.CONNECT
    SendFunction.cs         — DUCK.SEND
    QueryFunction.cs        — DUCK.QUERY
    ExecuteFunction.cs      — DUCK.EXECUTE
    InstructionsFunction.cs — DUCK.INSTRUCTIONS
```
