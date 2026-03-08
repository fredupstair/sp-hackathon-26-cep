using CepFunctions.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// Leaderboard API – public data (no per-app breakdown for other users).
/// GET /api/leaderboard?scope=global|department|team&amp;month=YYYY-MM&amp;page=1&amp;pageSize=20
/// </summary>
public class LeaderboardApi
{
    private readonly SharePointClient _sp;
    private readonly ILogger<LeaderboardApi> _log;

    public LeaderboardApi(SharePointClient sp, ILogger<LeaderboardApi> log)
    {
        _sp = sp;
        _log = log;
    }

    [Function("Leaderboard")]
    public async Task<IActionResult> GetAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "leaderboard")] HttpRequest req,
        CancellationToken ct)
    {
        var month = req.Query["month"].FirstOrDefault() ?? DateTime.UtcNow.ToString("yyyy-MM");
        var scopeParam = (req.Query["scope"].FirstOrDefault() ?? "global").ToLowerInvariant();
        int.TryParse(req.Query["page"].FirstOrDefault() ?? "1", out var page);
        int.TryParse(req.Query["pageSize"].FirstOrDefault() ?? "20", out var pageSize);
        if (page < 1) page = 1;
        if (pageSize is < 1 or > 100) pageSize = 20;

        // Resolve scope string.
        // For department/team we need the caller's profile (department/team name).
        string scope;
        if (scopeParam == "global")
        {
            scope = "Global";
        }
        else
        {
            // Caller must provide their AadUserId so we can look up their dept/team
            var callerId = req.Headers.TryGetValue("X-User-Id", out var v) ? v.FirstOrDefault() : null;
            if (callerId is null)
                return new BadRequestObjectResult("X-User-Id header required for department/team scope.");

            var caller = await _sp.GetUserByAadIdAsync(callerId, ct);
            if (caller is null)
                return new NotFoundObjectResult("Calling user not found.");

            scope = scopeParam == "department"
                ? $"Department:{caller.Department}"
                : $"Team:{caller.Team}";
        }

        var entries = await _sp.GetLeaderboardAsync(month, scope, page, pageSize, ct);
        var syncState = await _sp.GetSyncStateAsync(ct);

        return new OkObjectResult(new
        {
            month,
            scope,
            page,
            pageSize,
            lastUpdated = syncState.LastLeaderboardRebuildUtc?.ToUniversalTime().ToString("o"),
            entries = entries.Select(e => new
            {
                rank = e.Rank,
                displayName = e.DisplayName,
                department = e.Department,
                team = e.Team,
                monthlyPoints = e.MonthlyPoints,
                level = e.Level,
            }),
        });
    }
}
