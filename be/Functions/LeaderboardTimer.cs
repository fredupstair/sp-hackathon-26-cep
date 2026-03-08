using CepFunctions.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// Daily timer that rebuilds leaderboards, awards MonthlyMaster badges,
/// and sends leaderboard refresh notifications.
/// Schedule: every day at 08:00 UTC.
/// </summary>
public class LeaderboardTimer
{
    private readonly OrchestratorTimer _orchestrator;
    private readonly SharePointClient _sp;
    private readonly ILogger<LeaderboardTimer> _log;

    public LeaderboardTimer(OrchestratorTimer orchestrator, SharePointClient sp, ILogger<LeaderboardTimer> log)
    {
        _orchestrator = orchestrator;
        _sp = sp;
        _log = log;
    }

    [Function("LeaderboardTimer")]
    public async Task RunAsync(
        [TimerTrigger("0 0 8 * * *")] TimerInfo timer,
        CancellationToken ct)
    {
        var correlationId = $"LB-{Guid.NewGuid():N}";
        var runTime = DateTime.UtcNow;
        _log.LogInformation("[{CorrId}] LeaderboardTimer started at {Time}", correlationId, runTime);

        try
        {
            var config = await _sp.GetConfigAsync(ct);
            await _orchestrator.RebuildLeaderboardsAsync(correlationId, runTime, config, ct);

            var state = await _sp.GetSyncStateAsync(ct);
            state.LastLeaderboardRebuildUtc = runTime;
            await _sp.UpsertSyncStateAsync(state, ct);

            _log.LogInformation("[{CorrId}] LeaderboardTimer completed successfully", correlationId);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[{CorrId}] LeaderboardTimer failed", correlationId);
        }
    }
}
