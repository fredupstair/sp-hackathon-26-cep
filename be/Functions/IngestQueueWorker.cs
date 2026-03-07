using CepFunctions.Models;
using CepFunctions.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// Queue-triggered worker – processes one IngestMessage (one user per invocation).
/// 
/// Flow per message:
/// 1. Fetch Graph AI interactions in [watermark, runTime).
/// 2. Aggregate by (date, appKey).
/// 3. Upsert rows in CEP_ActivityLog (merge with existing day aggregates).
/// 4. Recalculate monthly points and level.
/// 5. Update CEP_Users (points, level, lastActivity, watermark).
/// 6. Evaluate & award badges, send Teams notifications.
/// </summary>
public class IngestQueueWorker
{
    private readonly GraphClient _graph;
    private readonly SharePointClient _sp;
    private readonly PointsEngine _points;
    private readonly BadgeEngine _badges;
    private readonly TeamsNotifier _notifier;
    private readonly IConfiguration _cfg;
    private readonly ILogger<IngestQueueWorker> _log;

    public IngestQueueWorker(
        GraphClient graph,
        SharePointClient sp,
        PointsEngine points,
        BadgeEngine badges,
        TeamsNotifier notifier,
        IConfiguration cfg,
        ILogger<IngestQueueWorker> log)
    {
        _graph = graph;
        _sp = sp;
        _points = points;
        _badges = badges;
        _notifier = notifier;
        _cfg = cfg;
        _log = log;
    }

    [Function("IngestQueueWorker")]
    public async Task RunAsync(
        [QueueTrigger("%IngestQueueName%")] IngestMessage msg,
        CancellationToken ct)
    {
        _log.LogInformation("[{CorrId}] Processing user {UserId} ({Upn})",
            msg.CorrelationId, msg.AadUserId, msg.UserPrincipalName);

        var config = await _sp.GetConfigAsync(ct);
        var monthKey = msg.RunTimeUtc.ToString("yyyy-MM");

        // ------------------------------------------------------------------
        // 1. Fetch Graph interactions and aggregate
        // ------------------------------------------------------------------
        Dictionary<(DateOnly Date, string AppKey), int> graphCounts;
        try
        {
            graphCounts = await _graph.GetDailyPromptCountsAsync(
                msg.AadUserId, msg.WatermarkUtc, msg.RunTimeUtc, ct);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[{CorrId}] Graph fetch failed for {UserId}", msg.CorrelationId, msg.AadUserId);
            throw; // Let the queue runtime retry (up to maxDequeueCount)
        }

        _log.LogInformation("[{CorrId}] User {UserId}: {Count} day/app buckets from Graph",
            msg.CorrelationId, msg.AadUserId, graphCounts.Count);

        // ------------------------------------------------------------------
        // 2. Load existing activity log for this month
        // ------------------------------------------------------------------
        var existingLogs = await _sp.GetActivityLogsForUserMonthAsync(msg.AadUserId, monthKey, ct);

        // ------------------------------------------------------------------
        // 3. Upsert CEP_ActivityLog rows
        // ------------------------------------------------------------------
        DateTime? latestActivity = null;
        foreach (var ((date, appKey), promptCount) in graphCounts)
        {
            var logKey = (appKey, date);
            var points = _points.ComputeDailyPoints(promptCount, config);

            if (existingLogs.TryGetValue(logKey, out var existing))
            {
                existing.PromptCount += promptCount;
                existing.PointsEarned += points;
                await _sp.UpsertActivityLogAsync(existing, ct);
            }
            else
            {
                var newLog = new CepActivityLog
                {
                    AadUserId = msg.AadUserId,
                    UserEmail = msg.UserEmail,
                    UsageDate = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Utc),
                    AppKey = appKey,
                    PromptCount = promptCount,
                    PointsEarned = points,
                    MonthKey = monthKey,
                };
                await _sp.UpsertActivityLogAsync(newLog, ct);
                existingLogs[logKey] = newLog;
            }

            var dt = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Utc);
            if (latestActivity is null || dt > latestActivity) latestActivity = dt;
        }

        // ------------------------------------------------------------------
        // 4. Recalculate user points & level
        // ------------------------------------------------------------------
        var user = await _sp.GetUserByAadIdAsync(msg.AadUserId, ct);
        if (user is null)
        {
            _log.LogWarning("[{CorrId}] User {UserId} not found in CEP_Users, skipping", msg.CorrelationId, msg.AadUserId);
            return;
        }

        string oldLevel = user.CurrentLevel;
        var allMonthLogs = await _sp.GetAllActivityLogsForUserMonthAsync(msg.AadUserId, monthKey, ct);
        (user.TotalPoints, user.MonthlyPoints) = _points.RecalculatePoints(user, allMonthLogs);
        user.CurrentLevel = _points.ComputeLevel(user.MonthlyPoints, config);

        if (latestActivity is not null && (user.LastActivityDate is null || latestActivity > user.LastActivityDate))
            user.LastActivityDate = latestActivity;

        user.LastSyncWatermarkUtc = msg.RunTimeUtc;

        // ------------------------------------------------------------------
        // 5. Evaluate badges
        // ------------------------------------------------------------------
        var existingBadges = await _sp.GetUserBadgesAsync(msg.AadUserId, ct);
        var newBadges = _badges.EvaluateBadges(
            user, allMonthLogs, existingBadges, config, monthKey, msg.RunTimeUtc).ToList();

        foreach (var badge in newBadges)
        {
            var awarded = await _sp.TryAwardBadgeAsync(badge, ct);
            if (awarded)
                await _notifier.SendBadgeEarnedAsync(user, badge, ct);
        }

        // ------------------------------------------------------------------
        // 6. Level-up notification (send only if level increased)
        // ------------------------------------------------------------------
        if (user.CurrentLevel != oldLevel)
            await _notifier.SendLevelUpAsync(user, user.CurrentLevel, ct);

        // ------------------------------------------------------------------
        // 7. Persist updated user
        // ------------------------------------------------------------------
        await _sp.UpsertUserAsync(user, ct);

        _log.LogInformation("[{CorrId}] User {UserId} processed. Points={Points}, Level={Level}, Badges={Badges}",
            msg.CorrelationId, msg.AadUserId, user.MonthlyPoints, user.CurrentLevel, newBadges.Count);
    }
}
