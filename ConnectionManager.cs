using DuckDB.NET.Data;
using System.Collections.Concurrent;

namespace DuckSheet;

public static class ConnectionManager
{
    private static readonly ConcurrentDictionary<string, string> _registry = new(StringComparer.OrdinalIgnoreCase);

    public static string Register(string path, string name)
    {
        // Verify the path is usable by opening and immediately closing a connection
        var conn = new DuckDBConnection($"Data Source={path}");
        conn.Open();
        conn.Dispose();

        _registry[name] = path;
        return $"Registered: {name} → {System.IO.Path.GetFileName(path)}  {DateTime.Now:M/d/yyyy HH:mm}";
    }

    public static DuckDBConnection Open(string name)
    {
        if (!_registry.TryGetValue(name, out var path))
            throw new InvalidOperationException($"No database registered as \"{name}\". Call DUCK.CONNECT first.");

        var conn = new DuckDBConnection($"Data Source={path}");
        conn.Open();
        return conn;
    }

    public static void Dispose()
    {
        _registry.Clear();
    }
}
