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
        if (v is JsonElement je)
        {
            // SP returns Number columns as JsonValueKind.Number; may be Double
            return je.ValueKind == JsonValueKind.Number
                ? (int)je.GetDouble()
                : int.TryParse(je.GetString(), out var pi) ? pi : fallback;
        }
        return int.TryParse(v.ToString(), out var i) ? i : fallback;
    }

    public static bool Bool(this Dictionary<string, object?> f, string key, bool fallback = false)
    {
        if (!f.TryGetValue(key, out var v) || v is null) return fallback;
        if (v is JsonElement je)
        {
            if (je.ValueKind == JsonValueKind.True) return true;
            if (je.ValueKind == JsonValueKind.False) return false;
            return bool.TryParse(je.GetString(), out var pb) ? pb : fallback;
        }
        return bool.TryParse(v.ToString(), out var b) ? b : fallback;
    }

    public static DateTime? Dt(this Dictionary<string, object?> f, string key)
    {
        if (!f.TryGetValue(key, out var v) || v is null) return null;
        var s = v is JsonElement je ? je.GetString() : v.ToString();
        return DateTime.TryParse(s, out var dt) ? dt : null;
    }
}
