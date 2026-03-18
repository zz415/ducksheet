using ExcelDna.Integration;

namespace DuckSheet;

public static class RangeWriter
{
    /// <summary>
    /// Schedules a write of data to a target range starting at the given ExcelReference.
    /// Must be called from a macro-type UDF context (IsMacroType=true).
    /// </summary>
    public static void WriteData(ExcelReference topLeft, object[,] data)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            var target = new ExcelReference(
                topLeft.RowFirst,
                topLeft.RowFirst + rows - 1,
                topLeft.ColumnFirst,
                topLeft.ColumnFirst + cols - 1,
                topLeft.SheetId);

            target.SetValue(data);
        });
    }

    /// <summary>
    /// Returns a display address string like "C20" for use in status returns.
    /// </summary>
    public static string GetAddress(ExcelReference reference)
    {
        try
        {
            return (string)XlCall.Excel(XlCall.xlfReftext, reference, true);
        }
        catch
        {
            return "?";
        }
    }
}
