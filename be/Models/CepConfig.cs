namespace CepFunctions.Models;

/// <summary>
/// Runtime configuration loaded from CEP_Config (key-value list).
/// Defaults are used when the list item is missing.
/// </summary>
public class CepConfig
{
    public string SyncFrequency { get; set; } = "daily";
    public int PointsPerPrompt { get; set; } = 1;
    public int LevelThresholdSilver { get; set; } = 500;
    public int LevelThresholdGold { get; set; } = 1500;
    public int InactivityDaysForNudge { get; set; } = 3;
    public bool LeaderboardRefreshNotificationEnabled { get; set; } = true;
    public int MaxUsersPerIngestionBatch { get; set; } = 50;
    public string TimeZone { get; set; } = "UTC";

    public static CepConfig FromSpItems(IEnumerable<Dictionary<string, object?>> items)
    {
        var cfg = new CepConfig();
        var lookup = items.ToDictionary(
            f => f.Str("Title"),
            f => f.Str("Value"));

        if (lookup.TryGetValue("SyncFrequency", out var v) && !string.IsNullOrEmpty(v)) cfg.SyncFrequency = v;
        if (lookup.TryGetValue("PointsPerPrompt", out v) && int.TryParse(v, out var i)) cfg.PointsPerPrompt = i;
        if (lookup.TryGetValue("LevelThresholdSilver", out v) && int.TryParse(v, out i)) cfg.LevelThresholdSilver = i;
        if (lookup.TryGetValue("LevelThresholdGold", out v) && int.TryParse(v, out i)) cfg.LevelThresholdGold = i;
        if (lookup.TryGetValue("InactivityDaysForNudge", out v) && int.TryParse(v, out i)) cfg.InactivityDaysForNudge = i;
        if (lookup.TryGetValue("LeaderboardRefreshNotificationEnabled", out v)) cfg.LeaderboardRefreshNotificationEnabled = v?.ToLower() == "true";
        if (lookup.TryGetValue("MaxUsersPerIngestionBatch", out v) && int.TryParse(v, out i)) cfg.MaxUsersPerIngestionBatch = i;
        if (lookup.TryGetValue("TimeZone", out v) && !string.IsNullOrEmpty(v)) cfg.TimeZone = v;

        return cfg;
    }

    public IEnumerable<Dictionary<string, object?>> ToSpRows() =>
    [
        Row("SyncFrequency", SyncFrequency),
        Row("PointsPerPrompt", PointsPerPrompt.ToString()),
        Row("LevelThresholdSilver", LevelThresholdSilver.ToString()),
        Row("LevelThresholdGold", LevelThresholdGold.ToString()),
        Row("InactivityDaysForNudge", InactivityDaysForNudge.ToString()),
        Row("LeaderboardRefreshNotificationEnabled", LeaderboardRefreshNotificationEnabled.ToString()),
        Row("MaxUsersPerIngestionBatch", MaxUsersPerIngestionBatch.ToString()),
        Row("TimeZone", TimeZone),
    ];

    private static Dictionary<string, object?> Row(string key, string value) =>
        new() { ["Title"] = key, ["Value"] = value };
}
