using CepFunctions.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// "Me" API – returns data scoped to the calling user.
///
/// Authentication:
///   - Production (EasyAuth enabled): caller identity is taken from the
///     X-MS-CLIENT-PRINCIPAL-ID header injected by Azure App Service Authentication.
///   - Local dev (EasyAuth not present): falls back to the X-User-Id header
///     so that local testing with REST clients still works.
///
/// GET /api/me/summary?month=YYYY-MM
/// GET /api/me/usage?from=YYYY-MM-DD&amp;to=YYYY-MM-DD
/// GET /api/me/badges
/// </summary>
public class MeApi
{
    private readonly SharePointClient _sp;
    private readonly ILogger<MeApi> _log;

    public MeApi(SharePointClient sp, ILogger<MeApi> log)
    {
        _sp = sp;
        _log = log;
    }

    // ------------------------------------------------------------------
    // GET /api/me/summary?month=YYYY-MM
    // ------------------------------------------------------------------
    [Function("MeSummary")]
    public async Task<IActionResult> GetSummaryAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "me/summary")] HttpRequest req,
        CancellationToken ct)
    {
        var aadUserId = GetCallerId(req);
        if (aadUserId is null) return Unauthorized();

        var user = await _sp.GetUserByAadIdAsync(aadUserId, ct);
        if (user is null) return new NotFoundObjectResult("User not enrolled.");

        var month = req.Query["month"].FirstOrDefault() ?? DateTime.UtcNow.ToString("yyyy-MM");

        // Fetch the user's global rank for the month
        var globalLb = await _sp.GetLeaderboardAsync(month, "Global", 1, int.MaxValue, ct);
        var entry = globalLb.FirstOrDefault(e => e.AadUserId == aadUserId);

        // Department / team rank
        var deptEntry = string.IsNullOrEmpty(user.Department) ? null
            : (await _sp.GetLeaderboardAsync(month, $"Department:{user.Department}", 1, int.MaxValue, ct))
              .FirstOrDefault(e => e.AadUserId == aadUserId);

        var teamEntry = string.IsNullOrEmpty(user.Team) ? null
            : (await _sp.GetLeaderboardAsync(month, $"Team:{user.Team}", 1, int.MaxValue, ct))
              .FirstOrDefault(e => e.AadUserId == aadUserId);

        return new OkObjectResult(new
        {
            displayName = user.DisplayName,
            currentLevel = user.CurrentLevel,
            totalPoints = user.TotalPoints,
            monthlyPoints = user.MonthlyPoints,
            month,
            globalRank = entry?.Rank,
            departmentRank = deptEntry?.Rank,
            teamRank = teamEntry?.Rank,
            lastActivityDate = user.LastActivityDate,
        });
    }

    // ------------------------------------------------------------------
    // GET /api/me/usage?from=YYYY-MM-DD&to=YYYY-MM-DD
    // ------------------------------------------------------------------
    [Function("MeUsage")]
    public async Task<IActionResult> GetUsageAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "me/usage")] HttpRequest req,
        CancellationToken ct)
    {
        var aadUserId = GetCallerId(req);
        if (aadUserId is null) return Unauthorized();

        var user = await _sp.GetUserByAadIdAsync(aadUserId, ct);
        if (user is null) return new NotFoundObjectResult("User not enrolled.");

        // Default to current month if no range specified
        var now = DateTime.UtcNow;
        var month = req.Query["month"].FirstOrDefault() ?? now.ToString("yyyy-MM");
        var logs = await _sp.GetActivityLogsForUserMonthAsync(aadUserId, month, ct);

        // Return breakdown by app
        var byApp = logs.Values
            .GroupBy(l => l.AppKey)
            .Select(g => new
            {
                appKey = g.Key,
                promptCount = g.Sum(l => l.PromptCount),
                pointsEarned = g.Sum(l => l.PointsEarned),
            })
            .OrderByDescending(x => x.promptCount);

        return new OkObjectResult(new { month, breakdown = byApp });
    }

    // ------------------------------------------------------------------
    // GET /api/me/badges
    // ------------------------------------------------------------------
    [Function("MeBadges")]
    public async Task<IActionResult> GetBadgesAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "me/badges")] HttpRequest req,
        CancellationToken ct)
    {
        var aadUserId = GetCallerId(req);
        if (aadUserId is null) return Unauthorized();

        var user = await _sp.GetUserByAadIdAsync(aadUserId, ct);
        if (user is null) return new NotFoundObjectResult("User not enrolled.");

        var badges = await _sp.GetUserBadgesAsync(aadUserId, ct);
        return new OkObjectResult(badges.Select(b => new
        {
            badgeKey = b.BadgeKey,
            badgeName = b.BadgeName,
            description = b.Description,
            earnedDate = b.EarnedDate,
            monthKey = b.MonthKey,
        }));
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    /// <summary>
    /// Returns the caller's Entra Object ID (OID).
    /// Production (EasyAuth enabled): taken from X-MS-CLIENT-PRINCIPAL-ID injected by Azure.
    /// Local dev (no EasyAuth): falls back to X-User-Id header for manual testing.
    /// </summary>
    private static string? GetCallerId(HttpRequest req)
    {
        if (req.Headers.TryGetValue("X-MS-CLIENT-PRINCIPAL-ID", out var easyAuth) &&
            easyAuth.FirstOrDefault() is { Length: > 0 } oid)
            return oid;

        // Local dev fallback – never trusted in production (EasyAuth blocks arbitrary headers)
        return req.Headers.TryGetValue("X-User-Id", out var devId) ? devId.FirstOrDefault() : null;
    }

    private static UnauthorizedObjectResult Unauthorized() =>
        new("Authenticated user identity not found. Ensure EasyAuth is enabled on the Function App.");
}
