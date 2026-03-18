using ExcelDna.Integration;

namespace DuckSheet.Functions;

public static class QueryFunction
{
    [ExcelFunction(Name = "DUCK.QUERY", Description = "Execute a SELECT and dump results starting at a target cell address.")]
    public static object Query(
        [ExcelArgument(Description = "SQL SELECT statement")] string sql,
        [ExcelArgument(Description = "Top-left cell address for output, e.g. \"E1\"")] string targetAddress,
        [ExcelArgument(Description = "Registered database name")] string dbName)
    {
        try
        {
            if (sql.EndsWith(".sql", StringComparison.OrdinalIgnoreCase) ||
                sql.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                sql = File.ReadAllText(sql).Trim();

            var rows = new List<object?[]>();
            string[]? headers = null;

            using (var conn = ConnectionManager.Open(dbName))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                using var reader = cmd.ExecuteReader();

                headers = new string[reader.FieldCount];
                for (int i = 0; i < reader.FieldCount; i++)
                    headers[i] = reader.GetName(i);

                while (reader.Read())
                {
                    var row = new object?[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                        row[i] = reader.IsDBNull(i) ? (object)ExcelEmpty.Value : ToExcelValue(reader.GetValue(i));
                    rows.Add(row);
                }
            }

            if (headers is null) return "Error: Query returned no schema.";

            int cols = headers.Length;
            var data = new object[rows.Count + 1, cols];
            for (int c = 0; c < cols; c++)
                data[0, c] = headers[c];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < cols; c++)
                    data[r + 1, c] = rows[r][c] ?? ExcelEmpty.Value;

            var capturedData = data;
            var capturedAddress = targetAddress;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    int nRows = capturedData.GetLength(0);
                    int nCols = capturedData.GetLength(1);

                    // xlSet cannot write cross-sheet — use COM instead
                    dynamic app = ExcelDnaUtil.Application;

                    // Parse optional sheet name from address e.g. "Results!$A$1"
                    string sheetName = null;
                    string cellAddr  = capturedAddress;
                    int bang = capturedAddress.IndexOf('!');
                    if (bang >= 0)
                    {
                        sheetName = capturedAddress.Substring(0, bang).Trim('\'');
                        cellAddr  = capturedAddress.Substring(bang + 1);
                    }

                    dynamic sheet = sheetName != null
                        ? app.ActiveWorkbook.Sheets[sheetName]
                        : app.ActiveSheet;

                    dynamic startCell = sheet.Range[cellAddr];
                    dynamic writeRange = sheet.Range[startCell, startCell.Offset[nRows - 1, nCols - 1]];

                    // Convert ExcelEmpty → null so COM writes blank cells
                    var comData = new object[nRows, nCols];
                    for (int r = 0; r < nRows; r++)
                        for (int c = 0; c < nCols; c++)
                        {
                            var v = capturedData[r, c];
                            comData[r, c] = v is ExcelEmpty ? null : v;
                        }

                    writeRange.Value = comData;
                }
                catch { /* swallow — never let a write failure crash Excel */ }
            });

            return $"Query OK: {rows.Count} rows → {targetAddress}  {DateTime.Now:M/d/yyyy HH:mm}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    // Convert DuckDB values to types Excel's SetValue actually accepts:
    // double, string, bool, DateTime, ExcelEmpty. Everything else crashes Excel.
    private static object ToExcelValue(object val) => val switch
    {
        double   => val,
        float  f => (double)f,
        bool     => val,
        string   => val,
        DateTime => val,
        int    i => (double)i,
        long   l => (double)l,
        short  s => (double)s,
        byte   b => (double)b,
        uint   u => (double)u,
        ulong  u => (double)u,
        ushort u => (double)u,
        decimal d => (double)d,
        ExcelEmpty => val,
        _ => (object?)val.ToString() ?? ExcelEmpty.Value
    };
}
