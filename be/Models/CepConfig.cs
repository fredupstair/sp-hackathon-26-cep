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
        // CEP_Config uses one row with named columns (not key-value rows)
        var cfg = new CepConfig();
        var first = items.FirstOrDefault();
        if (first is null) return cfg;

        if (first.Str("CEP_Cfg_SyncFrequency") is { Length: > 0 } sf) cfg.SyncFrequency = sf;
        if (int.TryParse(first.Str("CEP_Cfg_PointsPerPrompt"), out var ppp)) cfg.PointsPerPrompt = ppp;
        if (int.TryParse(first.Str("CEP_Cfg_LevelThresholdSilver"), out var s)) cfg.LevelThresholdSilver = s;
        if (int.TryParse(first.Str("CEP_Cfg_LevelThresholdGold"), out var g)) cfg.LevelThresholdGold = g;
        if (int.TryParse(first.Str("CEP_Cfg_InactivityDaysForNudge"), out var d)) cfg.InactivityDaysForNudge = d;
        cfg.LeaderboardRefreshNotificationEnabled = first.Bool("CEP_Cfg_LeaderboardRefreshNotifE", true);
        if (int.TryParse(first.Str("CEP_Cfg_MaxUsersPerIngestionBatc"), out var m)) cfg.MaxUsersPerIngestionBatch = m;
        if (first.Str("CEP_Cfg_TimeZone") is { Length: > 0 } tz) cfg.TimeZone = tz;

        return cfg;
    }

    public IEnumerable<Dictionary<string, object?>> ToSpRows() =>
    [
        new()
        {
            ["Title"] = "Config",
            ["CEP_Cfg_SyncFrequency"] = SyncFrequency,
            ["CEP_Cfg_PointsPerPrompt"] = PointsPerPrompt,
            ["CEP_Cfg_LevelThresholdSilver"] = LevelThresholdSilver,
            ["CEP_Cfg_LevelThresholdGold"] = LevelThresholdGold,
            ["CEP_Cfg_InactivityDaysForNudge"] = InactivityDaysForNudge,
            ["CEP_Cfg_LeaderboardRefreshNotifE"] = LeaderboardRefreshNotificationEnabled,
            ["CEP_Cfg_MaxUsersPerIngestionBatc"] = MaxUsersPerIngestionBatch,
            ["CEP_Cfg_TimeZone"] = TimeZone,
        }
    ];
}
