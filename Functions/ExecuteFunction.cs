using ExcelDna.Integration;

namespace DuckSheet.Functions;

public static class ExecuteFunction
{
    [ExcelFunction(Name = "DUCK.EXECUTE", Description = "Execute a SQL statement that returns no rows (DDL, DML, PRAGMA).")]
    public static object Execute(
        [ExcelArgument(Description = "SQL statement to execute")] string sql,
        [ExcelArgument(Description = "Registered database name")] string dbName)
    {
        try
        {
            if (sql.EndsWith(".sql", StringComparison.OrdinalIgnoreCase) ||
                sql.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                sql = File.ReadAllText(sql).Trim();

            using var conn = ConnectionManager.Open(dbName);
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            int affected = cmd.ExecuteNonQuery();
            return $"{DescribeSql(sql, affected)}  {DateTime.Now:M/d/yyyy HH:mm}";
        }
        catch (Exception ex)
        {
            return $"Error: {ex.Message}";
        }
    }

    private static string DescribeSql(string sql, int affected)
    {
        var trimmed = sql.Trim();
        var tokens = trimmed.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0) return "OK";

        string verb = tokens[0].ToUpperInvariant();
        string sub = tokens.Length > 1 ? tokens[1].ToUpperInvariant() : "";
        string name = tokens.Length > 2 ? tokens[2].Trim('"', '\'', '`', '[', ']') : "";

        return (verb, sub) switch
        {
            ("CREATE", "TABLE")    => $"Created table: {name}",
            ("CREATE", "VIEW")     => $"Created view: {name}",
            ("CREATE", "INDEX")    => $"Created index: {name}",
            ("CREATE", "SEQUENCE") => $"Created sequence: {name}",
            ("CREATE", "SCHEMA")   => $"Created schema: {name}",
            ("CREATE", "MACRO")    => $"Created macro: {name}",
            ("CREATE", "OR")       => tokens.Length > 4
                                        ? $"Created or replaced {tokens[3].ToLower()}: {tokens[4].Trim('"', '\'', '`', '[', ']')}"
                                        : "Created or replaced",
            ("DROP", "TABLE")      => $"Dropped table: {name}",
            ("DROP", "VIEW")       => $"Dropped view: {name}",
            ("DROP", "INDEX")      => $"Dropped index: {name}",
            ("DROP", "SEQUENCE")   => $"Dropped sequence: {name}",
            ("DROP", "SCHEMA")     => $"Dropped schema: {name}",
            ("DROP", "MACRO")      => $"Dropped macro: {name}",
            ("ALTER", "TABLE")     => $"Altered table: {name}",
            ("INSERT", _)          => affected >= 0 ? $"Inserted: {affected} rows" : "Insert OK",
            ("UPDATE", _)          => affected >= 0 ? $"Updated: {affected} rows" : "Update OK",
            ("DELETE", _)          => affected >= 0 ? $"Deleted: {affected} rows" : "Delete OK",
            ("TRUNCATE", _)        => $"Truncated: {name}",
            ("PRAGMA", _)          => $"Pragma: {sub.ToLower()}",
            ("SET", _)             => $"Set: {sub.ToLower()}",
            _                      => "OK"
        };
    }
}
