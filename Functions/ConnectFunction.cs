using ExcelDna.Integration;

namespace DuckSheet.Functions;

public static class ConnectFunction
{
    [ExcelFunction(Name = "DUCK.CONNECT", Description = "Register a DuckDB file under a name for use in other DUCK functions.")]
    public static object Connect(
        [ExcelArgument(Description = "Full path to the .duckdb file")] string path,
        [ExcelArgument(Description = "Name to register this database as, e.g. \"db1\"")] string name)
    {
        try
        {
            return ConnectionManager.Register(path, name);
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }
}
