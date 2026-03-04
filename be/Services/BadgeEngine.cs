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
            yield return NewBadge(user, CepBadge.Keys.FirstSteps, "First Steps",
                "Executed your first Copilot prompt after enrollment.", now, "");
        }

        // ------------------------------------------------------------------
        // CrossAppExplorer – 3+ distinct apps in any calendar week
        // ------------------------------------------------------------------
        if (!awarded.Contains(CepBadge.Keys.CrossAppExplorer))
        {
            var qualifies = logs
                .GroupBy(l => IsoWeekKey(l.UsageDate))
                .Any(g => g.Select(l => l.AppKey).Distinct().Count() >= 3);

            if (qualifies)
                yield return NewBadge(user, CepBadge.Keys.CrossAppExplorer, "Cross-App Explorer",
                    "Used 3 or more Copilot apps in a single week.", now, monthKey);
        }

        // ------------------------------------------------------------------
        // WeeklyWarrior – 50+ prompts in any calendar week
        // ------------------------------------------------------------------
        if (!awarded.Contains(CepBadge.Keys.WeeklyWarrior))
        {
            var qualifies = logs
                .GroupBy(l => IsoWeekKey(l.UsageDate))
                .Any(g => g.Sum(l => l.PromptCount) >= 50);

            if (qualifies)
                yield return NewBadge(user, CepBadge.Keys.WeeklyWarrior, "Weekly Warrior",
                    "Completed 50 or more prompts in a single week.", now, monthKey);
        }

        // ------------------------------------------------------------------
        // ConsistencyKing – 7 consecutive active days (any 7-day window)
        // ------------------------------------------------------------------
        if (!awarded.Contains(CepBadge.Keys.ConsistencyKing))
        {
            var activeDays = logs
                .Where(l => l.PromptCount > 0)
                .Select(l => DateOnly.FromDateTime(l.UsageDate.Date))
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            if (HasConsecutiveDays(activeDays, 7))
                yield return NewBadge(user, CepBadge.Keys.ConsistencyKing, "Consistency King",
                    "Active for 7 consecutive days.", now, monthKey);
        }

        // Note: MonthlyMaster (Top 10 this month) is awarded after leaderboard calculation,
        // not here. See OrchestratorTimer.
    }

    /// <summary>Awards MonthlyMaster badge to users in top 10 of the global leaderboard.</summary>
    public IEnumerable<(CepUser User, CepBadge Badge)> EvaluateMonthlyMaster(
        List<CepLeaderboardEntry> globalLeaderboard,
        IEnumerable<CepUser> allUsers,
        string monthKey,
        DateTime now)
    {
        var top10 = globalLeaderboard.Where(e => e.Rank <= 10).Select(e => e.AadUserId).ToHashSet();
        foreach (var user in allUsers.Where(u => top10.Contains(u.AadUserId)))
        {
            yield return (user, NewBadge(user, CepBadge.Keys.MonthlyMaster, "Monthly Master",
                "Ranked in the Top 10 of the monthly leaderboard.", now, monthKey));
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
