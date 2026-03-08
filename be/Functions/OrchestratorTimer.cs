using System.Text.Json;
using Azure.Storage.Queues;
using CepFunctions.Models;
using CepFunctions.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// Timer-triggered orchestrator.
/// Schedule: fires every minute. Each invocation reads CEP_Config.SyncIntervalMinutes
/// and CEP_SyncState.LastSuccessfulRunUtc to decide whether a full sync is due.
/// This way, changing the interval from SharePoint takes effect within one minute.
/// 
/// When a sync IS due:
/// 1. Load active users.
/// 2. Enqueue one IngestMessage per user → Queue Worker fan-out.
/// 3. Rebuild leaderboards and send refresh notifications.
/// 4. Write run result to CEP_SyncState.
/// </summary>
public class OrchestratorTimer
{
    private readonly SharePointClient _sp;
    private readonly PointsEngine _points;
    private readonly BadgeEngine _badges;
    private readonly TeamsNotifier _notifier;
    private readonly IConfiguration _cfg;
    private readonly ILogger<OrchestratorTimer> _log;

    public OrchestratorTimer(
        SharePointClient sp,
        PointsEngine points,
        BadgeEngine badges,
        TeamsNotifier notifier,
        IConfiguration cfg,
        ILogger<OrchestratorTimer> log)
    {
        _sp = sp;
        _points = points;
        _badges = badges;
        _notifier = notifier;
        _cfg = cfg;
        _log = log;
    }

    [Function("OrchestratorTimer")]
    public async Task RunAsync(
        [TimerTrigger("0 */1 * * * *")] TimerInfo timer,
        CancellationToken ct)
    {
        // Every minute: read config + sync state to decide if a full sync is due.
        var config = await _sp.GetConfigAsync(ct);
        var state = await _sp.GetSyncStateAsync(ct);
        var now = DateTime.UtcNow;

        var minutesSinceLast = state.LastSuccessfulRunUtc.HasValue
            ? (now - state.LastSuccessfulRunUtc.Value).TotalMinutes
            : double.MaxValue; // never ran → force run

        if (minutesSinceLast < config.SyncIntervalMinutes)
        {
            _log.LogDebug("Sync not due yet ({Elapsed:F0} min < {Interval} min). Skipping.",
                minutesSinceLast, config.SyncIntervalMinutes);
            return;
        }

        // Guard against overlapping runs
        if (state.LastRunStatus == "Running")
        {
            _log.LogWarning("Previous sync is still running (correlation {CorrId}). Skipping.",
                state.LastRunCorrelationId);
            return;
        }

        await RunSyncAsync(ct);
    }

    /// <summary>Core sync logic – callable from the timer trigger or from the admin HTTP endpoint.</summary>
    public async Task<(int Enqueued, string Status)> RunSyncAsync(CancellationToken ct)
    {
        var correlationId = Guid.NewGuid().ToString("N");
        var runTime = DateTime.UtcNow;
        _log.LogInformation("[{CorrId}] Orchestrator started at {Time}", correlationId, runTime);

        var state = await _sp.GetSyncStateAsync(ct);
        state.LastRunStatus = "Running";
        state.LastRunCorrelationId = correlationId;
        await _sp.UpsertSyncStateAsync(state, ct);

        int enqueued = 0;
        try
        {
            var config = await _sp.GetConfigAsync(ct);
            var users = await _sp.GetActiveUsersAsync(ct);
            _log.LogInformation("[{CorrId}] Enqueuing {Count} users", correlationId, users.Count);

            var queueName = _cfg["IngestQueueName"] ?? "cep-ingest";
            var queueClient = new QueueClient(_cfg["AzureWebJobsStorage"], queueName);
            await queueClient.CreateIfNotExistsAsync(cancellationToken: ct);

            foreach (var user in users.Take(config.MaxUsersPerIngestionBatch))
            {
                var msg = new IngestMessage
                {
                    AadUserId = user.AadUserId,
                    UserPrincipalName = user.UserPrincipalName,
                    UserEmail = user.Email,
                    WatermarkUtc = user.LastSyncWatermarkUtc ?? runTime.AddDays(-1),
                    RunTimeUtc = runTime,
                    CorrelationId = correlationId,
                };
                var json = JsonSerializer.Serialize(msg);
                var encoded = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(json));
                await queueClient.SendMessageAsync(encoded, cancellationToken: ct);
                enqueued++;
            }

            _log.LogInformation("[{CorrId}] Enqueued {Count} messages", correlationId, enqueued);

            await SendInactivityNudgesAsync(correlationId, config, runTime, ct);

            state.LastRunStatus = "Success";
            state.LastSuccessfulRunUtc = runTime;
            state.LastRunSummary = $"Enqueued {enqueued} users for ingest.";
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[{CorrId}] Orchestrator failed", correlationId);
            state.LastRunStatus = "Failure";
            state.LastRunSummary = ex.Message[..Math.Min(500, ex.Message.Length)];
        }
        finally
        {
            await _sp.UpsertSyncStateAsync(state, ct);
        }
        return (enqueued, state.LastRunStatus);
    }

    // ------------------------------------------------------------------
    // Leaderboard
    // ------------------------------------------------------------------

    /// <summary>Rebuilds leaderboards, awards MonthlyMaster badges, and sends notifications. Called by LeaderboardTimer.</summary>
    public async Task RebuildLeaderboardsAsync(string corrId, DateTime runTime, CepConfig config, CancellationToken ct)
    {
        var monthKey = runTime.ToString("yyyy-MM");
        var users = await _sp.GetActiveUsersAsync(ct);
        _log.LogInformation("[{CorrId}] Rebuilding leaderboards for {Month}", corrId, monthKey);

        foreach (var (scope, entries) in _points.BuildAllLeaderboards(users, monthKey, config))
        {
            await _sp.ReplaceLeaderboardAsync(monthKey, scope, entries, ct);
        }

        // MonthlyMaster badge
        var global = await _sp.GetLeaderboardAsync(monthKey, "Global", 1, 10, ct);
        foreach (var (user, badge) in _badges.EvaluateMonthlyMaster(global, users, monthKey, runTime, config))
        {
            var awarded = await _sp.TryAwardBadgeAsync(badge, ct);
            if (awarded)
                await _notifier.SendBadgeEarnedAsync(user, badge, config, ct);
        }

        // Leaderboard refresh notifications
        if (config.LeaderboardRefreshNotificationEnabled)
        {
            _log.LogInformation("[{CorrId}] LeaderboardRefreshNotification enabled – checking {Count} users for month {Month}",
                corrId, users.Count, monthKey);

            foreach (var user in users)
            {
                if (user.LastLeaderboardNotifiedMonth == monthKey)
                {
                    _log.LogInformation("[{CorrId}] User {Upn} already notified for {Month}, skipping",
                        corrId, user.UserPrincipalName, monthKey);
                    continue;
                }

                var entry = global.FirstOrDefault(e => e.AadUserId == user.AadUserId);
                // Skip users with no points (rank is meaningless)
                if (entry is null || entry.MonthlyPoints <= 0)
                {
                    _log.LogInformation("[{CorrId}] User {Upn} has no points for {Month}, skipping leaderboard notification",
                        corrId, user.UserPrincipalName, monthKey);
                    continue;
                }

                _log.LogInformation("[{CorrId}] Sending leaderboard notification to {Upn} (rank={Rank}, month={Month})",
                    corrId, user.UserPrincipalName, entry.Rank, monthKey);

                await _notifier.SendLeaderboardUpdateAsync(user, entry.Rank, monthKey, config, ct);

                user.LastLeaderboardNotifiedMonth = monthKey;
                await _sp.UpsertUserAsync(user, ct);
            }
        }
    }

    // ------------------------------------------------------------------
    // Inactivity nudges
    // ------------------------------------------------------------------

    private async Task SendInactivityNudgesAsync(string corrId, CepConfig config, DateTime now, CancellationToken ct)
    {
        var users = await _sp.GetActiveUsersAsync(ct);
        _log.LogInformation("[{CorrId}] Checking inactivity nudges for {Count} users", corrId, users.Count);

        foreach (var user in users)
        {
            var sent = await _notifier.SendInactivityNudgeIfDueAsync(user, config, now, ct);
            if (sent)
            {
                user.LastNudgeSentUtc = now;
                await _sp.UpsertUserAsync(user, ct);
            }
        }
    }
}
