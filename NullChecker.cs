using ExcelDna.Integration;

namespace DuckSheet;

public static class NullChecker
{
    public static bool IsNull(object? cellValue)
    {
        if (cellValue is null) return true;
        if (cellValue is DBNull) return true;
        if (cellValue is ExcelEmpty) return true;
        if (cellValue is ExcelMissing) return true;
        if (cellValue is ExcelError) return true;
        if (cellValue is string s && s.Trim() == "") return true;
        var str = cellValue.ToString();
        if (str == "NULL" || str == "") return true;
        return false;
    }
}
