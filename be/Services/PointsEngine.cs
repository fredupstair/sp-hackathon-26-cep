using CepFunctions.Models;

namespace CepFunctions.Services;

/// <summary>
/// Calculates points, determines level and builds leaderboard rankings.
/// Stateless – all inputs are passed explicitly.
/// </summary>
public class PointsEngine
{
    // ------------------------------------------------------------------
    // Points
    // ------------------------------------------------------------------

    /// <summary>Returns points earned for a single day's prompt count.</summary>
    public int ComputeDailyPoints(int promptCount, CepConfig cfg) =>
        promptCount * cfg.PointsPerPrompt;

    /// <summary>
    /// Merges new daily aggregates into the existing monthly totals of a user.
    /// Returns updated TotalPoints and MonthlyPoints.
    /// </summary>
    public (int TotalPoints, int MonthlyPoints) RecalculatePoints(
        CepUser user,
        IEnumerable<CepActivityLog> monthLogs)
    {
        var monthly = monthLogs.Sum(l => l.PointsEarned);
        // TotalPoints = keep everything already accumulated in previous months  
        // minus the old monthly figure, then add new monthly.
        var total = user.TotalPoints - user.MonthlyPoints + monthly;
        return (total, monthly);
    }

    // ------------------------------------------------------------------
    // Levels
    // ------------------------------------------------------------------

    public string ComputeLevel(int monthlyPoints, CepConfig cfg)
    {
        if (monthlyPoints >= cfg.LevelThresholdGold) return "Gold";
        if (monthlyPoints >= cfg.LevelThresholdSilver) return "Silver";
        return "Bronze";
    }

    // ------------------------------------------------------------------
    // Leaderboard materialisation
    // ------------------------------------------------------------------

    /// <summary>
    /// Builds ranked leaderboard entries from a flat list of users for a given month/scope.
    /// Uses standard competition ranking (same score → same rank, next rank skips).
    /// </summary>
    public List<CepLeaderboardEntry> BuildLeaderboard(
        IEnumerable<CepUser> users,
        string monthKey,
        string scope,
        CepConfig cfg)
    {
        var sorted = users
            .OrderByDescending(u => u.MonthlyPoints)
            .ThenBy(u => u.DisplayName)
            .ToList();

        var entries = new List<CepLeaderboardEntry>(sorted.Count);
        int rank = 1;
        int lastPoints = -1;
        int lastRank = 1;

        for (int i = 0; i < sorted.Count; i++)
        {
            var u = sorted[i];
            if (u.MonthlyPoints != lastPoints)
            {
                lastRank = rank;
                lastPoints = u.MonthlyPoints;
            }

            entries.Add(new CepLeaderboardEntry
            {
                MonthKey = monthKey,
                Scope = scope,
                AadUserId = u.AadUserId,
                UserEmail = u.Email,
                DisplayName = u.DisplayName,
                Department = u.Department,
                Team = u.Team,
                MonthlyPoints = u.MonthlyPoints,
                Rank = lastRank,
                Level = ComputeLevel(u.MonthlyPoints, cfg),
            });

            rank++;
        }

        return entries;
    }

    /// <summary>
    /// Builds Global + per-Department + per-Team leaderboards in one pass.
    /// </summary>
    public IEnumerable<(string Scope, List<CepLeaderboardEntry> Entries)> BuildAllLeaderboards(
        IEnumerable<CepUser> allUsers,
        string monthKey,
        CepConfig cfg)
    {
        var users = allUsers.Where(u => u.IsActive).ToList();

        yield return ("Global", BuildLeaderboard(users, monthKey, "Global", cfg));

        foreach (var dept in users.Select(u => u.Department).Where(d => !string.IsNullOrEmpty(d)).Distinct())
        {
            var deptUsers = users.Where(u => u.Department == dept).ToList();
            yield return ($"Department:{dept}", BuildLeaderboard(deptUsers, monthKey, $"Department:{dept}", cfg));
        }

        foreach (var team in users.Select(u => u.Team).Where(t => !string.IsNullOrEmpty(t)).Distinct())
        {
            var teamUsers = users.Where(u => u.Team == team).ToList();
            yield return ($"Team:{team}", BuildLeaderboard(teamUsers, monthKey, $"Team:{team}", cfg));
        }
    }
}
