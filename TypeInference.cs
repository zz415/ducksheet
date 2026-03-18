namespace DuckSheet;

public enum DuckType { BOOLEAN, TIMESTAMP, BIGINT, DOUBLE, VARCHAR }

public static class TypeInference
{
    // Excel serial date range: 1 (1900-01-01) to 2958465 (9999-12-31)
    private const double ExcelDateMin = 1.0;
    private const double ExcelDateMax = 2958465.0;

    /// <summary>
    /// Infers the narrowest DuckDB type for a column of values.
    /// Returns the inferred type and the number of values coerced to NULL.
    /// </summary>
    public static (DuckType Type, int NullCount) InferColumn(object?[] values, bool forceTimestamp = false)
    {
        var nonNull = new List<object>();
        int nullCount = 0;

        foreach (var v in values)
        {
            if (NullChecker.IsNull(v))
                nullCount++;
            else
                nonNull.Add(v!);
        }

        if (nonNull.Count == 0)
            return (DuckType.VARCHAR, nullCount);

        if (forceTimestamp)
            return (DuckType.TIMESTAMP, nullCount);

        // Try from narrowest to widest
        if (AllBoolean(nonNull)) return (DuckType.BOOLEAN, nullCount);
        if (AllTimestamp(nonNull)) return (DuckType.TIMESTAMP, nullCount);
        if (AllBigInt(nonNull)) return (DuckType.BIGINT, nullCount);
        if (AllDouble(nonNull)) return (DuckType.DOUBLE, nullCount);
        return (DuckType.VARCHAR, nullCount);
    }

    private static bool AllBoolean(List<object> values)
    {
        foreach (var v in values)
        {
            if (v is bool) continue;
            if (v is string s)
            {
                if (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                    s.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                    continue;
            }
            return false;
        }
        return true;
    }

    private static bool AllTimestamp(List<object> values)
    {
        // Only infer TIMESTAMP from string values that parse as dates.
        // Pure numeric columns stay BIGINT/DOUBLE to avoid Excel serial ambiguity.
        foreach (var v in values)
        {
            if (v is string s)
            {
                if (DateTime.TryParse(s, out _)) continue;
            }
            // Doubles and other types do NOT qualify as TIMESTAMP in V1
            return false;
        }
        return true;
    }

    private static bool AllBigInt(List<object> values)
    {
        foreach (var v in values)
        {
            if (v is double d)
            {
                if (d == Math.Floor(d) && !double.IsInfinity(d) && !double.IsNaN(d))
                    continue;
                return false;
            }
            if (v is long or int) continue;
            if (v is string s && long.TryParse(s, out _)) continue;
            return false;
        }
        return true;
    }

    private static bool AllDouble(List<object> values)
    {
        foreach (var v in values)
        {
            if (v is double or float) continue;
            if (v is long or int) continue;
            if (v is string s && double.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out _)) continue;
            return false;
        }
        return true;
    }

    public static string ToDuckSql(DuckType type) => type switch
    {
        DuckType.BOOLEAN => "BOOLEAN",
        DuckType.TIMESTAMP => "TIMESTAMP",
        DuckType.BIGINT => "BIGINT",
        DuckType.DOUBLE => "DOUBLE",
        _ => "VARCHAR",
    };
}
