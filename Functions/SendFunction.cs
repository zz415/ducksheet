using ExcelDna.Integration;
using System.Text;

namespace DuckSheet.Functions;

public static class SendFunction
{
    [ExcelFunction(Name = "DUCK.SEND", Description = "Push an Excel range into a DuckDB table (DROP + recreate).", IsMacroType = true)]
    public static object Send(
        [ExcelArgument(Description = "Range to push (first row = headers)", AllowReference = true)] object rangeArg,
        [ExcelArgument(Description = "Target table name in DuckDB")] string tableName,
        [ExcelArgument(Description = "Registered database name")] string dbName)
    {
        try
        {
            var rangeRef = rangeArg as ExcelReference;
            object[,]? range = null;
            if (rangeRef != null)
            {
                var v = rangeRef.GetValue();
                if (v is object[,] arr) range = arr;
            }
            else
            {
                range = rangeArg as object[,];
            }

            if (range is null || range.Length == 0)
                return "Error: Range is empty.";
            if (string.IsNullOrWhiteSpace(tableName))
                return "Error: Table name is required.";
            if (string.IsNullOrWhiteSpace(dbName))
                return "Error: Database name is required.";

            int rows = range.GetLength(0);
            int cols = range.GetLength(1);

            if (rows < 2)
                return "Error: Range must have at least a header row and one data row.";

            var headers = new string[cols];
            for (int c = 0; c < cols; c++)
            {
                var h = range[0, c]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(h)) h = $"col{c + 1}";
                headers[c] = SanitizeIdentifier(h);
            }

            int dataRows = rows - 1;
            var columns = new object?[cols][];
            for (int c = 0; c < cols; c++)
            {
                columns[c] = new object?[dataRows];
                for (int r = 0; r < dataRows; r++)
                    columns[c][r] = range[r + 1, c];
            }

            // Detect Excel-formatted date columns by checking the number format
            // of the first non-null double in each column.
            var forceTimestamp = new bool[cols];
            if (rangeRef != null)
            {
                for (int c = 0; c < cols; c++)
                {
                    for (int r = 0; r < dataRows; r++)
                    {
                        var val = columns[c][r];
                        if (val is double && !NullChecker.IsNull(val))
                        {
                            var cellRef = new ExcelReference(
                                rangeRef.RowFirst + r + 1,
                                rangeRef.RowFirst + r + 1,
                                rangeRef.ColumnFirst + c,
                                rangeRef.ColumnFirst + c,
                                rangeRef.SheetId);
                            var fmt = XlCall.Excel(XlCall.xlfGetCell, 7, cellRef) as string;
                            if (fmt != null && IsDateFormat(fmt))
                                forceTimestamp[c] = true;
                            break;
                        }
                    }
                }
            }

            var types = new DuckType[cols];
            var nullCounts = new int[cols];
            for (int c = 0; c < cols; c++)
                (types[c], nullCounts[c]) = TypeInference.InferColumn(columns[c], forceTimestamp[c]);

            var safeTable = SanitizeIdentifier(tableName);

            using (var conn = ConnectionManager.Open(dbName))
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = $"DROP TABLE IF EXISTS \"{safeTable}\"";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = BuildDdl(safeTable, headers, types);
                    cmd.ExecuteNonQuery();
                }

                using var appender = conn.CreateAppender(safeTable);
                for (int r = 0; r < dataRows; r++)
                {
                    var row = appender.CreateRow();
                    for (int c = 0; c < cols; c++)
                    {
                        var val = columns[c][r];
                        if (NullChecker.IsNull(val))
                        { row.AppendNullValue(); continue; }
                        AppendTyped(row, val!, types[c]);
                    }
                    row.EndRow();
                }
            }

            var sb = new StringBuilder();
            sb.AppendLine($"Loaded: {safeTable} — {dataRows} rows, {cols} cols");
            for (int c = 0; c < cols; c++)
            {
                var nullNote = nullCounts[c] > 0 ? $" ({nullCounts[c]} nulls)" : "";
                sb.AppendLine($"  {headers[c]}: {TypeInference.ToDuckSql(types[c])}{nullNote}");
            }
            return sb.ToString().TrimEnd() + $"  {DateTime.Now:M/d/yyyy HH:mm}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    private static string BuildDdl(string tableName, string[] headers, DuckType[] types)
    {
        var cols = headers.Select((h, i) => $"\"{h}\" {TypeInference.ToDuckSql(types[i])}");
        return $"CREATE TABLE \"{tableName}\" ({string.Join(", ", cols)})";
    }

    private static void AppendTyped(DuckDB.NET.Data.IDuckDBAppenderRow row, object val, DuckType type)
    {
        try
        {
            switch (type)
            {
                case DuckType.BOOLEAN:
                    if (val is bool b) { row.AppendValue(b); return; }
                    row.AppendValue(bool.Parse(val.ToString()!));
                    return;
                case DuckType.TIMESTAMP:
                    if (val is double serial) { row.AppendValue(DateTime.FromOADate(serial)); return; }
                    if (val is string s && DateTime.TryParse(s, out var dt))
                    { row.AppendValue(dt); return; }
                    row.AppendNullValue();
                    return;
                case DuckType.BIGINT:
                    if (val is double d) { row.AppendValue((long)d); return; }
                    row.AppendValue(long.Parse(val.ToString()!));
                    return;
                case DuckType.DOUBLE:
                    if (val is double dbl) { row.AppendValue(dbl); return; }
                    row.AppendValue(double.Parse(val.ToString()!,
                        System.Globalization.CultureInfo.InvariantCulture));
                    return;
                default:
                    row.AppendValue(val.ToString());
                    return;
            }
        }
        catch { row.AppendNullValue(); }
    }

    // Date formats always contain 'y' (year). Time-only and numeric formats don't.
    private static bool IsDateFormat(string fmt) =>
        fmt.IndexOf('y', StringComparison.OrdinalIgnoreCase) >= 0;

    private static string SanitizeIdentifier(string name)
    {
        var sb = new StringBuilder();
        foreach (var ch in name.Trim())
            sb.Append(char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_');
        var result = sb.ToString();
        if (result.Length > 0 && char.IsDigit(result[0]))
            result = "_" + result;
        return result.Length > 0 ? result : "col";
    }
}
