namespace CepFunctions.Models;

/// <summary>
/// Badge definition loaded from CEP_Config key-value rows (category = "badges").
/// </summary>
public record BadgeDefinition(string Key, string Name, string Description, int Threshold = 0);

/// <summary>
/// Runtime configuration loaded from CEP_Config (key-value list with categories).
/// Each SP row has: Title (key), CEP_Cfg_Value (value), CEP_Cfg_Category (category).
/// Defaults are used when the list item is missing.
/// </summary>
public class CepConfig
{
    // ── Sync ──────────────────────────────────────────────────
    public int SyncIntervalMinutes { get; set; } = 60;
    public int MaxUsersPerIngestionBatch { get; set; } = 50;

    // ── Scoring ───────────────────────────────────────────────
    public int PointsPerPrompt { get; set; } = 1;
    public int PointsPerWin { get; set; } = 10;
    public int MaxWinsPerDay { get; set; } = 10;
    public int LevelThresholdSilver { get; set; } = 500;
    public int LevelThresholdGold { get; set; } = 1500;

    // ── General ───────────────────────────────────────────────
    public int InactivityDaysForNudge { get; set; } = 3;
    public bool LeaderboardRefreshNotificationEnabled { get; set; } = true;
    public string TimeZone { get; set; } = "UTC";

    // ── Badges ────────────────────────────────────────────────
    public List<BadgeDefinition> BadgeDefinitions { get; set; } =
    [
        new("FirstSteps",       "First Steps",       "Executed your first Copilot prompt after enrollment."),
        new("CrossAppExplorer", "Cross-App Explorer", "Used 3+ Copilot apps in a single week.", 3),
        new("WeeklyWarrior",    "Weekly Warrior",     "50+ prompts in a single week.", 50),
        new("ConsistencyKing",  "Consistency King",   "Active for 7 consecutive days.", 7),
        new("MonthlyMaster",    "Monthly Master",     "Top 10 in the monthly leaderboard.", 10),
    ];

    /// <summary>Shortcut to look up a badge definition by key.</summary>
    public BadgeDefinition GetBadge(string key) =>
        BadgeDefinitions.FirstOrDefault(b => b.Key == key)
        ?? new BadgeDefinition(key, key, "", 0);

    // ------------------------------------------------------------------
    // Serialisation from/to SharePoint key-value rows
    // ------------------------------------------------------------------

    public static CepConfig FromSpItems(IEnumerable<Dictionary<string, object?>> items)
    {
        var cfg = new CepConfig();
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var row in items)
        {
            var key = row.Str("Title");
            var val = row.Str("CEP_Cfg_Value");
            if (!string.IsNullOrEmpty(key))
                map[key] = val;
        }

        // Sync
        if (map.TryGetValue("SyncIntervalMinutes", out var sim) && int.TryParse(sim, out var simV)) cfg.SyncIntervalMinutes = simV;
        if (map.TryGetValue("MaxUsersPerIngestionBatch", out var mub) && int.TryParse(mub, out var mubV)) cfg.MaxUsersPerIngestionBatch = mubV;

        // Scoring
        if (map.TryGetValue("PointsPerPrompt", out var ppp) && int.TryParse(ppp, out var pppV)) cfg.PointsPerPrompt = pppV;
        if (map.TryGetValue("PointsPerWin", out var ppw) && int.TryParse(ppw, out var ppwV)) cfg.PointsPerWin = ppwV;
        if (map.TryGetValue("MaxWinsPerDay", out var mwd) && int.TryParse(mwd, out var mwdV)) cfg.MaxWinsPerDay = mwdV;
        if (map.TryGetValue("LevelThresholdSilver", out var lts) && int.TryParse(lts, out var ltsV)) cfg.LevelThresholdSilver = ltsV;
        if (map.TryGetValue("LevelThresholdGold", out var ltg) && int.TryParse(ltg, out var ltgV)) cfg.LevelThresholdGold = ltgV;

        // General
        if (map.TryGetValue("InactivityDaysForNudge", out var idn) && int.TryParse(idn, out var idnV)) cfg.InactivityDaysForNudge = idnV;
        if (map.TryGetValue("LeaderboardRefreshNotificationEnabled", out var lrne))
            cfg.LeaderboardRefreshNotificationEnabled = bool.TryParse(lrne, out var lrneV) && lrneV;
        if (map.TryGetValue("TimeZone", out var tz) && !string.IsNullOrEmpty(tz)) cfg.TimeZone = tz;

        // Badges – override defaults from key-value rows
        var badges = new List<BadgeDefinition>(cfg.BadgeDefinitions);
        for (int i = 0; i < badges.Count; i++)
        {
            var b = badges[i];
            var name = map.TryGetValue($"Badge_{b.Key}_Name", out var bn) && !string.IsNullOrEmpty(bn) ? bn : b.Name;
            var desc = map.TryGetValue($"Badge_{b.Key}_Description", out var bd) && !string.IsNullOrEmpty(bd) ? bd : b.Description;
            var thr = map.TryGetValue($"Badge_{b.Key}_Threshold", out var bt) && int.TryParse(bt, out var btV) ? btV : b.Threshold;
            badges[i] = new BadgeDefinition(b.Key, name, desc, thr);
        }
        cfg.BadgeDefinitions = badges;

        return cfg;
    }

    public IEnumerable<Dictionary<string, object?>> ToSpRows()
    {
        var rows = new List<(string Key, string Value, string Category)>
        {
            // Sync
            ("SyncIntervalMinutes", SyncIntervalMinutes.ToString(), "sync"),
            ("MaxUsersPerIngestionBatch", MaxUsersPerIngestionBatch.ToString(), "sync"),
            // Scoring
            ("PointsPerPrompt", PointsPerPrompt.ToString(), "scoring"),
            ("PointsPerWin", PointsPerWin.ToString(), "scoring"),
            ("MaxWinsPerDay", MaxWinsPerDay.ToString(), "scoring"),
            ("LevelThresholdSilver", LevelThresholdSilver.ToString(), "scoring"),
            ("LevelThresholdGold", LevelThresholdGold.ToString(), "scoring"),
            // General
            ("InactivityDaysForNudge", InactivityDaysForNudge.ToString(), "general"),
            ("LeaderboardRefreshNotificationEnabled", LeaderboardRefreshNotificationEnabled.ToString(), "general"),
            ("TimeZone", TimeZone, "general"),
        };

        // Badges
        foreach (var b in BadgeDefinitions)
        {
            rows.Add(($"Badge_{b.Key}_Name", b.Name, "badges"));
            rows.Add(($"Badge_{b.Key}_Description", b.Description, "badges"));
            if (b.Threshold > 0)
                rows.Add(($"Badge_{b.Key}_Threshold", b.Threshold.ToString(), "badges"));
        }

        return rows.Select(r => new Dictionary<string, object?>
        {
            ["Title"] = r.Key,
            ["CEP_Cfg_Value"] = r.Value,
            ["CEP_Cfg_Category"] = r.Category,
        });
    }
}
