using ExcelDna.Integration;

namespace DuckSheet.Functions;

public static class InstructionsFunction
{
    [ExcelFunction(Name = "DUCK.INSTRUCTIONS", Description = "Returns usage instructions for all DuckSheet functions.")]
    public static string Instructions()
    {
        return """
=== DUCKSHEET INSTRUCTIONS ===

-- STEP 1: Register a database --
=DUCK.CONNECT("C:\path\to\file.duckdb", "db1")
  Creates the .duckdb file if it doesn't exist, or connects to an existing one.
  "db1" is the name you use in all other functions to refer to this database.
  You can register multiple databases:
    =DUCK.CONNECT("C:\data\sales.duckdb", "sales")
    =DUCK.CONNECT("C:\data\hr.duckdb", "hr")
  Returns: Registered: db1 → file.duckdb

-- STEP 2: Push a range to DuckDB --
=DUCK.SEND(A1:F100, "table_name", "db1")
  First row of the range must be column headers.
  Drops and recreates the table on every call (overwrite semantics).
  Types are inferred per column: BOOLEAN → TIMESTAMP → BIGINT → DOUBLE → VARCHAR
  Empty cells, errors, and blanks become NULL.
  Returns: Loaded: table_name — 99 rows, 6 cols
             col1: VARCHAR
             col2: BIGINT (2 nulls)
             ...

-- STEP 3: Run a SELECT query --
=DUCK.QUERY("SELECT * FROM table_name", "E1", "db1")
  Results are written starting at the target cell (headers in row 1).
  Always pass the target as a plain string address — do NOT use CELL().
  CELL() is a volatile function and will cause an infinite recalculation loop.
  Returns: Query OK: 99 rows → E1

  Address examples:
    "E1"              -- cell on the current sheet
    "$E$1"            -- absolute reference, same sheet
    "Results!A1"      -- cell on another sheet named Results
    "Results!$A$1"    -- absolute reference on another sheet

  You can pass a .sql or .txt file path instead of writing SQL inline:
  =DUCK.QUERY("C:\queries\monthly_report.sql", "Results!$A$1", "db1")
  =DUCK.QUERY("C:\queries\my_query.txt", "E1", "db1")

-- STEP 4: Run DDL or DML --
=DUCK.EXECUTE("CREATE VIEW v AS SELECT ...", "db1")
  Use for: CREATE TABLE, CREATE VIEW, DROP, INSERT, ALTER, PRAGMA.
  Use DUCK.QUERY for anything that returns rows, DUCK.EXECUTE for everything else.
  Returns: action description + timestamp, e.g. "Created view: v  3/17/2026 19:38"

  You can pass a .sql or .txt file path instead of writing SQL inline:
  =DUCK.EXECUTE("C:\queries\build_views.sql", "db1")
  =DUCK.EXECUTE("C:\queries\setup.txt", "db1")

-- DESCRIBE a table schema --
=DUCK.QUERY("DESCRIBE table_name", "H1", "db1")
  DuckDB's native DESCRIBE returns column name, type, and nullability.

-- CONTROL PANEL LAYOUT (recommended) --
  Place all DUCK formulas in a dedicated sheet (e.g. "DuckDB").
  Excel evaluates top-to-bottom so CONNECT must be above SEND and QUERY.

  B2:  =DUCK.CONNECT("C:\data\mywork.duckdb", "db1")
  B4:  =DUCK.SEND(Data!A1:F500, "sales", "db1")
  B5:  =DUCK.SEND(Data!H1:K200, "products", "db1")
  B7:  =DUCK.EXECUTE("CREATE OR REPLACE VIEW summary AS SELECT ...", "db1")
  B9:  =DUCK.QUERY("SELECT * FROM summary", "Results!$A$1", "db1")

-- DBEAVER CONNECTION (read-only) --
  Driver:    DuckDB
  File path: same path used in DUCK.CONNECT
  Property:  access_mode = read_only
  Each DUCK function opens and closes its connection per call,
  so DBeaver can connect freely between operations.

-- REFRESH --
  Single formula: click the cell → F2 → Enter
  Full refresh:   Ctrl+Alt+F9 (recalculates all formulas top to bottom)
""";
    }
}
