using System.Text.Json;

namespace CepFunctions.Models;

/// <summary>
/// Extension helpers for reading SharePoint field dictionaries safely.
/// </summary>
public static class SpFieldExtensions
{
    public static string Str(this Dictionary<string, object?> f, string key, string fallback = "") =>
        f.TryGetValue(key, out var v) ? v?.ToString() ?? fallback : fallback;

    public static int Int(this Dictionary<string, object?> f, string key, int fallback = 0)
    {
        if (!f.TryGetValue(key, out var v) || v is null) return fallback;
        return v is JsonElement je ? je.GetInt32() : int.TryParse(v.ToString(), out var i) ? i : fallback;
    }

    public static bool Bool(this Dictionary<string, object?> f, string key, bool fallback = false)
    {
        if (!f.TryGetValue(key, out var v) || v is null) return fallback;
        if (v is JsonElement je) return je.GetBoolean();
        return bool.TryParse(v.ToString(), out var b) ? b : fallback;
    }

    public static DateTime? Dt(this Dictionary<string, object?> f, string key)
    {
        if (!f.TryGetValue(key, out var v) || v is null) return null;
        var s = v is JsonElement je ? je.GetString() : v.ToString();
        return DateTime.TryParse(s, out var dt) ? dt : null;
    }
}
