using System.IO;
using System.Text.Json;
using CepFunctions.Models;
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
/// POST /api/me/wins
/// GET /api/me/suggestion
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

        var now = DateTime.UtcNow;
        var month = req.Query["month"].FirstOrDefault() ?? now.ToString("yyyy-MM");

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

        // Compute recent active days (last 2 months, excludes Win entries) for client-side streak calculation
        var currMonthKey = now.ToString("yyyy-MM");
        var prevMonthKey = now.AddMonths(-1).ToString("yyyy-MM");
        var optInDate = user.EnrollmentDate.HasValue
            ? DateOnly.FromDateTime(user.EnrollmentDate.Value.ToUniversalTime())
            : (DateOnly?)null;
        var currLogs = await _sp.GetAllActivityLogsForUserMonthAsync(aadUserId, currMonthKey, ct);
        var prevLogs = await _sp.GetAllActivityLogsForUserMonthAsync(aadUserId, prevMonthKey, ct);
        var recentActiveDays = currLogs.Concat(prevLogs)
            .Where(l => !optInDate.HasValue || DateOnly.FromDateTime(l.UsageDate) >= optInDate.Value)
            .Where(l => !l.IsWin && l.PromptCount > 0)
            .Select(l => l.UsageDate.ToString("yyyy-MM-dd"))
            .Distinct()
            .OrderByDescending(d => d)
            .Take(60)
            .ToArray();

        // For the current month use the live values from the user record (always up to date).
        // For past months use the values materialised in the leaderboard entry for that month,
        // so that the month selector shows historically correct points and level.
        bool isCurrentMonth = month == currMonthKey;
        int displayMonthlyPoints = isCurrentMonth ? user.MonthlyPoints : (entry?.MonthlyPoints ?? 0);
        string displayLevel      = isCurrentMonth ? user.CurrentLevel  : (entry?.Level ?? user.CurrentLevel);

        return new OkObjectResult(new
        {
            userId = user.AadUserId,
            displayName = user.DisplayName,
            email = user.Email,
            department = user.Department,
            team = user.Team,
            enrollmentDate = user.EnrollmentDate,
            isActive = user.IsActive,
            isEngagementNudgesEnabled = user.IsEngagementNudgesEnabled,
            currentLevel = displayLevel,
            totalPoints = user.TotalPoints,
            monthlyPoints = displayMonthlyPoints,
            month,
            globalRank = entry?.Rank,
            departmentRank = deptEntry?.Rank,
            teamRank = teamEntry?.Rank,
            lastActivityDate = user.LastActivityDate,
            recentActiveDays,
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

        // Accept ?from=YYYY-MM-DD&to=YYYY-MM-DD; derive month from the from-date.
        // Fall back to the current month when parameters are missing.
        var now = DateTime.UtcNow;
        var fromStr = req.Query["from"].FirstOrDefault();
        var toStr   = req.Query["to"].FirstOrDefault();

        DateTime fromDate = DateTime.TryParse(fromStr, out var fd) ? fd : new DateTime(now.Year, now.Month, 1);
        DateTime toDate   = DateTime.TryParse(toStr,   out var td) ? td : new DateTime(now.Year, now.Month,
            DateTime.DaysInMonth(now.Year, now.Month));
        var optInDate = user.EnrollmentDate.HasValue
            ? DateOnly.FromDateTime(user.EnrollmentDate.Value.ToUniversalTime())
            : (DateOnly?)null;

        var month = fromDate.ToString("yyyy-MM");
        var logs = await _sp.GetAllActivityLogsForUserMonthAsync(aadUserId, month, ct);
        var endExclusive = toDate.Date.AddDays(1);
        logs = logs
            .Where(l => l.UsageDate >= fromDate.Date && l.UsageDate < endExclusive)
            .Where(l => !optInDate.HasValue || DateOnly.FromDateTime(l.UsageDate) >= optInDate.Value)
            .ToList();

        // Return breakdown by app + totals
        var breakdown = logs
            .GroupBy(l => l.AppKey)
            .Select(g => new
            {
                appKey = g.Key,
                promptCount = g.Sum(l => l.PromptCount),
                pointsEarned = g.Sum(l => l.PointsEarned),
            })
            .OrderByDescending(x => x.promptCount)
            .ToList();

        return new OkObjectResult(new
        {
            from = fromDate.ToString("yyyy-MM-dd"),
            to   = toDate.ToString("yyyy-MM-dd"),
            breakdown,
            totalPrompts = breakdown.Sum(x => x.promptCount),
            totalPoints  = breakdown.Sum(x => x.pointsEarned),
        });
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
    // Static data for suggestion engine
    // ------------------------------------------------------------------

    private static readonly string[] _suggestableApps =
        ["Word", "Excel", "PowerPoint", "Outlook", "Teams", "OneNote", "Loop", "BizChat", "WebChat", "M365App", "Forms", "SharePoint", "Whiteboard"];

    private static readonly Dictionary<string, (string Label, string Text)> _suggestionPool = new()
    {
        ["Word"]       = ("Word",               "Draft your next document with Copilot in Word — just describe what you need and let it write the first version!"),
        ["Excel"]      = ("Excel",              "Ask Copilot in Excel to analyse your data or build a formula — save hours on manual work!"),
        ["PowerPoint"] = ("PowerPoint",         "Create a whole presentation from a single prompt with Copilot in PowerPoint!"),
        ["Outlook"]    = ("Outlook",            "Summarise long email threads in seconds with Copilot in Outlook!"),
        ["Teams"]      = ("Teams",              "Catch up on missed meetings: ask Copilot in Teams to summarise the last transcript!"),
        ["OneNote"]    = ("OneNote",            "Turn scattered notes into tidy action items with Copilot in OneNote!"),
        ["Loop"]       = ("Loop",               "Brainstorm in real-time with Copilot in Loop — great for async collaboration!"),
        ["BizChat"]    = ("Copilot Chat",       "Ask Copilot Chat a question about any of your M365 files — it searches across all your content!"),
        ["WebChat"]    = ("Web Chat",           "Use Copilot via the web — no desktop app needed!"),
        ["M365App"]    = ("M365 App",           "Explore Copilot features across the wider Microsoft 365 ecosystem!"),
        ["Forms"]      = ("Forms",              "Let Copilot help you draft surveys and quizzes in Microsoft Forms!"),
        ["SharePoint"] = ("SharePoint",         "Use Copilot in SharePoint to design pages and surfaces faster!"),
        ["Whiteboard"] = ("Whiteboard",         "Ideate visually with Copilot in Whiteboard — turn brainstorms into structured plans!"),
    };

    private static readonly JsonSerializerOptions _jsonOpts = new() { PropertyNameCaseInsensitive = true };

    private record WinRequest(string AppKey, string? Note, bool IsShared);

    // ------------------------------------------------------------------
    // POST /api/me/wins
    // ------------------------------------------------------------------

    [Function("MePostWin")]
    public async Task<IActionResult> PostWinAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "me/wins")] HttpRequest req,
        CancellationToken ct)
    {
        var aadUserId = GetCallerId(req);
        if (aadUserId is null) return Unauthorized();

        var user = await _sp.GetUserByAadIdAsync(aadUserId, ct);
        if (user is null) return new NotFoundObjectResult("User not enrolled.");

        using var reader = new StreamReader(req.Body);
        var body = await reader.ReadToEndAsync();
        WinRequest? winReq;
        try { winReq = JsonSerializer.Deserialize<WinRequest>(body, _jsonOpts); }
        catch { return new BadRequestObjectResult("Invalid request body."); }
        if (winReq is null || string.IsNullOrWhiteSpace(winReq.AppKey))
            return new BadRequestObjectResult("appKey is required.");

        var config = await _sp.GetConfigAsync(ct);
        var now = DateTime.UtcNow;
        var today = DateOnly.FromDateTime(now);
        var monthKey = now.ToString("yyyy-MM");

        // Rate limit: count individual win rows for today
        var todayCount = await _sp.CountWinsForUserDayAsync(aadUserId, today, ct);

        if (todayCount >= config.MaxWinsPerDay)
            return new ObjectResult(new { error = "Daily win limit reached.", todayWinCount = todayCount })
                   { StatusCode = 429 };

        // Each win → its own row (preserves note, app context, and future metadata)
        var winLog = new CepActivityLog
        {
            AadUserId = aadUserId,
            UserEmail = user.Email,
            UsageDate = now,
            AppKey = "Win",
            PromptCount = 1,
            PointsEarned = config.PointsPerWin,
            MonthKey = monthKey,
            IsWin = true,
            WinNote = winReq.Note ?? "",
            IsShared = winReq.IsShared,
            WinAppKey = winReq.AppKey,
        };

        await _sp.UpsertActivityLogAsync(winLog, ct);

        user.MonthlyPoints += config.PointsPerWin;
        user.TotalPoints += config.PointsPerWin;
        await _sp.UpsertUserAsync(user, ct);

        return new OkObjectResult(new
        {
            pointsAdded = config.PointsPerWin,
            totalMonthlyPoints = user.MonthlyPoints,
            todayWinCount = todayCount + 1,
        });
    }

    // ------------------------------------------------------------------
    // GET /api/me/suggestion
    // ------------------------------------------------------------------

    [Function("MeSuggestion")]
    public async Task<IActionResult> GetSuggestionAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "me/suggestion")] HttpRequest req,
        CancellationToken ct)
    {
        var aadUserId = GetCallerId(req);
        if (aadUserId is null) return Unauthorized();

        var user = await _sp.GetUserByAadIdAsync(aadUserId, ct);
        if (user is null) return new NotFoundObjectResult("User not enrolled.");

        var monthKey = DateTime.UtcNow.ToString("yyyy-MM");
        var logs = await _sp.GetAllActivityLogsForUserMonthAsync(aadUserId, monthKey, ct);

        var usageByApp = logs
            .Where(l => !l.IsWin)
            .GroupBy(l => l.AppKey)
            .ToDictionary(g => g.Key, g => g.Sum(l => l.PromptCount));

        // Pick the suggestable app with the fewest prompts this month; random tiebreak
        var picked = _suggestableApps
            .Select(k => (appKey: k, count: usageByApp.TryGetValue(k, out var c) ? c : 0))
            .OrderBy(x => x.count)
            .ThenBy(_ => Guid.NewGuid())
            .First();

        if (!_suggestionPool.TryGetValue(picked.appKey, out var suggestion))
            return new NoContentResult();

        return new OkObjectResult(new
        {
            appKey = picked.appKey,
            appLabel = suggestion.Label,
            text = suggestion.Text,
        });
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
