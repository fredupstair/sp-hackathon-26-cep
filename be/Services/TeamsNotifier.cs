using CepFunctions.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Services;

/// <summary>
/// Sends Teams activity notifications using GraphClient.
/// Contains all idempotency and anti-spam logic.
/// </summary>
public class TeamsNotifier
{
    private readonly GraphClient _graph;
    private readonly string _teamsAppId;
    private readonly ILogger<TeamsNotifier> _log;

    public TeamsNotifier(GraphClient graph, IConfiguration cfg, ILogger<TeamsNotifier> log)
    {
        _graph = graph;
        _teamsAppId = cfg["TeamsAppId"] ?? "";
        _log = log;
    }

    // ------------------------------------------------------------------
    // Welcome
    // ------------------------------------------------------------------

    public Task SendWelcomeAsync(CepUser user, CepConfig cfg, CancellationToken ct = default) =>
        TrySendAsync(user.UserPrincipalName, "cepWelcome",
            "Welcome to the Copilot Engagement Program! Start earning points by using Copilot. \U0001F680",
            null, cfg.NotificationUrlDashboard, ct);

    // ------------------------------------------------------------------
    // Level Up
    // ------------------------------------------------------------------

    public Task SendLevelUpAsync(CepUser user, string newLevel, CepConfig cfg, CancellationToken ct = default) =>
        TrySendAsync(user.UserPrincipalName, "cepLevelUp",
            $"Congratulations {user.DisplayName}! You reached {newLevel} level. Keep it up! \U0001F3C6",
            new() { ["levelName"] = newLevel }, cfg.NotificationUrlDashboard, ct);

    // ------------------------------------------------------------------
    // Badge Earned
    // ------------------------------------------------------------------

    public Task SendBadgeEarnedAsync(CepUser user, CepBadge badge, CepConfig cfg, CancellationToken ct = default) =>
        TrySendAsync(user.UserPrincipalName, "cepBadgeEarned",
            $"You earned the '{badge.BadgeName}' badge! {badge.Description} \U0001F396\uFE0F",
            new() { ["badgeName"] = badge.BadgeName, ["badgeDescription"] = badge.Description }, cfg.NotificationUrlDashboard, ct);

    // ------------------------------------------------------------------
    // Inactivity Nudge (anti-spam: max 1 every 3 days)
    // ------------------------------------------------------------------

    /// <summary>Returns true if the nudge was sent and LastNudgeSentUtc should be updated.</summary>
    public async Task<bool> SendInactivityNudgeIfDueAsync(CepUser user, CepConfig cfg, DateTime now, CancellationToken ct = default)
    {
        if (!user.IsEngagementNudgesEnabled) return false;
        if (user.LastActivityDate is not null &&
            (now - user.LastActivityDate.Value).TotalDays < cfg.InactivityDaysForNudge)
            return false;

        // Anti-spam cooldown: wait at least InactivityDaysForNudge between nudges
        if (user.LastNudgeSentUtc is not null &&
            (now - user.LastNudgeSentUtc.Value).TotalDays < cfg.InactivityDaysForNudge)
            return false;

        await TrySendAsync(user.UserPrincipalName, "cepInactivityNudge",
            $"Hey {user.DisplayName}, we miss you! Open Copilot and earn some points. \U0001F4A1",
            new() { ["userName"] = user.DisplayName }, cfg.NotificationUrlCopilotChat, ct);
        return true;
    }

    // ------------------------------------------------------------------
    // Leaderboard Refresh
    // ------------------------------------------------------------------

    public Task SendLeaderboardUpdateAsync(CepUser user, int rank, string monthKey, CepConfig cfg, CancellationToken ct = default) =>
        TrySendAsync(user.UserPrincipalName, "cepLeaderboardUpdate",
            $"Leaderboard updated for {monthKey}. Your current global rank: #{rank}. Keep going! \U0001F4CA",
            new() { ["monthKey"] = monthKey, ["rank"] = rank.ToString() }, cfg.NotificationUrlLeaderboard, ct);

    // ------------------------------------------------------------------
    // Internal
    // ------------------------------------------------------------------

    private async Task TrySendAsync(string upn, string activityType, string preview,
        Dictionary<string, string>? templateParams, string? webUrl, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(upn))
        {
            _log.LogWarning("Cannot send Teams notification – empty UPN");
            return;
        }
        try
        {
            await _graph.SendActivityNotificationAsync(upn, _teamsAppId, activityType, preview, templateParams, webUrl, ct);
            _log.LogInformation("Teams notification '{Type}' sent to {Upn}", activityType, upn);
        }
        catch (Exception ex)
        {
            // Notifications are best-effort – never fail the main flow
            _log.LogWarning(ex, "Teams notification '{Type}' failed for {Upn}", activityType, upn);
        }
    }
}
