using CepFunctions.Models;

namespace CepFunctions.Services;

/// <summary>
/// Evaluates badge criteria against activity aggregates and user state.
/// Returns the list of badges to award (not yet persisted).
/// </summary>
public class BadgeEngine
{
    /// <summary>
    /// Returns zero or more new badges based on current activity logs and user state.
    /// Already-awarded badges must be filtered by the caller using TryAwardBadgeAsync.
    /// Badge metadata (names, descriptions, thresholds) are driven by CepConfig.
    /// </summary>
    public IEnumerable<CepBadge> EvaluateBadges(
        CepUser user,
        IEnumerable<CepActivityLog> allLogs,
        IEnumerable<CepBadge> existingBadges,
        CepConfig cfg,
        string monthKey,
        DateTime now)
    {
        var logs = allLogs.ToList();
        var awarded = new HashSet<string>(existingBadges.Select(b => b.BadgeKey));

        // ------------------------------------------------------------------
        // FirstSteps – first prompt ever after enrollment
        // ------------------------------------------------------------------
        if (!awarded.Contains(CepBadge.Keys.FirstSteps) && logs.Sum(l => l.PromptCount) > 0)
        {
            var def = cfg.GetBadge(CepBadge.Keys.FirstSteps);
            yield return NewBadge(user, def.Key, def.Name, def.Description, now, "");
        }

        // ------------------------------------------------------------------
        // CrossAppExplorer – N+ distinct apps in any calendar week
        // ------------------------------------------------------------------
        if (!awarded.Contains(CepBadge.Keys.CrossAppExplorer))
        {
            var def = cfg.GetBadge(CepBadge.Keys.CrossAppExplorer);
            var threshold = def.Threshold > 0 ? def.Threshold : 3;
            var qualifies = logs
                .Where(l => !l.IsWin)
                .GroupBy(l => IsoWeekKey(l.UsageDate))
                .Any(g => g.Select(l => l.AppKey).Distinct().Count() >= threshold);

            if (qualifies)
                yield return NewBadge(user, def.Key, def.Name, def.Description, now, monthKey);
        }

        // ------------------------------------------------------------------
        // WeeklyWarrior – N+ prompts in any calendar week
        // ------------------------------------------------------------------
        if (!awarded.Contains(CepBadge.Keys.WeeklyWarrior))
        {
            var def = cfg.GetBadge(CepBadge.Keys.WeeklyWarrior);
            var threshold = def.Threshold > 0 ? def.Threshold : 50;
            var qualifies = logs
                .GroupBy(l => IsoWeekKey(l.UsageDate))
                .Any(g => g.Sum(l => l.PromptCount) >= threshold);

            if (qualifies)
                yield return NewBadge(user, def.Key, def.Name, def.Description, now, monthKey);
        }

        // ------------------------------------------------------------------
        // ConsistencyKing – N consecutive active days (any N-day window)
        // ------------------------------------------------------------------
        if (!awarded.Contains(CepBadge.Keys.ConsistencyKing))
        {
            var def = cfg.GetBadge(CepBadge.Keys.ConsistencyKing);
            var threshold = def.Threshold > 0 ? def.Threshold : 7;
            var activeDays = logs
                .Where(l => l.PromptCount > 0)
                .Select(l => DateOnly.FromDateTime(l.UsageDate.Date))
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            if (HasConsecutiveDays(activeDays, threshold))
                yield return NewBadge(user, def.Key, def.Name, def.Description, now, monthKey);
        }

        // Note: MonthlyMaster (Top N this month) is awarded after leaderboard calculation,
        // not here. See OrchestratorTimer.
    }

    /// <summary>Awards MonthlyMaster badge to users in top N of the global leaderboard.</summary>
    public IEnumerable<(CepUser User, CepBadge Badge)> EvaluateMonthlyMaster(
        List<CepLeaderboardEntry> globalLeaderboard,
        IEnumerable<CepUser> allUsers,
        string monthKey,
        DateTime now,
        CepConfig cfg)
    {
        var def = cfg.GetBadge(CepBadge.Keys.MonthlyMaster);
        var threshold = def.Threshold > 0 ? def.Threshold : 10;
        var topN = globalLeaderboard.Where(e => e.Rank <= threshold).Select(e => e.AadUserId).ToHashSet();
        foreach (var user in allUsers.Where(u => topN.Contains(u.AadUserId)))
        {
            yield return (user, NewBadge(user, def.Key, def.Name, def.Description, now, monthKey));
        }
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    private static CepBadge NewBadge(CepUser user, string key, string name, string description, DateTime now, string monthKey) =>
        new()
        {
            AadUserId = user.AadUserId,
            UserEmail = user.Email,
            BadgeKey = key,
            BadgeName = name,
            Description = description,
            EarnedDate = now,
            MonthKey = monthKey,
        };

    private static string IsoWeekKey(DateTime dt)
    {
        var day = (int)dt.DayOfWeek;
        var startOfWeek = dt.AddDays(-(day == 0 ? 6 : day - 1)).Date;
        return startOfWeek.ToString("yyyy-MM-dd");
    }

    private static bool HasConsecutiveDays(List<DateOnly> sortedDays, int required)
    {
        if (sortedDays.Count < required) return false;
        int streak = 1;
        for (int i = 1; i < sortedDays.Count; i++)
        {
            streak = sortedDays[i] == sortedDays[i - 1].AddDays(1) ? streak + 1 : 1;
            if (streak >= required) return true;
        }
        return false;
    }
}
