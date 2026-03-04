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
/// Schedule: every day at 03:00 UTC (configurable via CRON in host.json / env).
/// 
/// Responsibilities:
/// 1. Load CEP_Config and active users.
/// 2. Enqueue one IngestMessage per user → Queue Worker fan-out.
/// 3. After ingest window (same run, after a delay or in a separate scheduled run),
///    rebuild leaderboards and send refresh notifications.
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
        [TimerTrigger("0 0 3 * * *")] TimerInfo timer,
        CancellationToken ct)
    {
        var correlationId = Guid.NewGuid().ToString("N");
        var runTime = DateTime.UtcNow;
        _log.LogInformation("[{CorrId}] Orchestrator started at {Time}", correlationId, runTime);

        var state = await _sp.GetSyncStateAsync(ct);
        state.LastRunStatus = "Running";
        state.LastRunCorrelationId = correlationId;
        await _sp.UpsertSyncStateAsync(state, ct);

        try
        {
            var config = await _sp.GetConfigAsync(ct);
            var users = await _sp.GetActiveUsersAsync(ct);
            _log.LogInformation("[{CorrId}] Enqueuing {Count} users", correlationId, users.Count);

            var queueName = _cfg["IngestQueueName"] ?? "cep-ingest";
            var queueClient = new QueueClient(_cfg["AzureWebJobsStorage"], queueName);
            await queueClient.CreateIfNotExistsAsync(cancellationToken: ct);

            int enqueued = 0;
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

            // -----------------------------------------------------------
            // Leaderboard rebuild (runs after ingest in this same timer OR
            // you can move this to a separate daily timer offset by +2 hours)
            // -----------------------------------------------------------
            await RebuildLeaderboardsAsync(correlationId, runTime, config, ct);

            // Inactivity nudges
            await SendInactivityNudgesAsync(correlationId, config, runTime, ct);

            // Persist success
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
    }

    // ------------------------------------------------------------------
    // Leaderboard
    // ------------------------------------------------------------------

    private async Task RebuildLeaderboardsAsync(string corrId, DateTime runTime, CepConfig config, CancellationToken ct)
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
        foreach (var (user, badge) in _badges.EvaluateMonthlyMaster(global, users, monthKey, runTime))
        {
            var awarded = await _sp.TryAwardBadgeAsync(badge, ct);
            if (awarded)
                await _notifier.SendBadgeEarnedAsync(user, badge, ct);
        }

        // Leaderboard refresh notifications
        if (config.LeaderboardRefreshNotificationEnabled)
        {
            foreach (var user in users)
            {
                if (user.LastLeaderboardNotifiedMonth == monthKey) continue;

                var rank = global.FirstOrDefault(e => e.AadUserId == user.AadUserId)?.Rank ?? 0;
                await _notifier.SendLeaderboardUpdateAsync(user, rank, monthKey, ct);

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
