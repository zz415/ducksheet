using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace DuckSheet;

public class AddIn : IExcelAddIn
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool AddDllDirectory(string lpPathName);

    public void AutoOpen()
    {
        // Add the XLL's directory to the native DLL search path so
        // DuckDB.NET can find duckdb.dll placed alongside the XLL.
        var xllDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        if (!string.IsNullOrEmpty(xllDir))
            AddDllDirectory(xllDir);
    }

    public void AutoClose()
    {
        ConnectionManager.Dispose();
    }
}
