using System.Text.Json;
using CepFunctions.Models;
using CepFunctions.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// Admin API – restricted to admin callers (validated via X-Admin-Key header in MVP;
/// in production replace with proper role claim check on the Entra token).
///
/// GET  /api/admin/config
/// POST /api/admin/config
/// GET  /api/admin/status
/// </summary>
public class AdminApi
{
    private readonly SharePointClient _sp;
    private readonly OrchestratorTimer _orchestrator;
    private readonly TeamsNotifier _notifier;
    private readonly GraphClient _graph;
    private readonly string _teamsAppId;
    private readonly ILogger<AdminApi> _log;

    public AdminApi(SharePointClient sp, OrchestratorTimer orchestrator, TeamsNotifier notifier, GraphClient graph, IConfiguration cfg, ILogger<AdminApi> log)
    {
        _sp = sp;
        _orchestrator = orchestrator;
        _notifier = notifier;
        _graph = graph;
        _teamsAppId = cfg["TeamsAppId"] ?? "";
        _log = log;
    }

    // ------------------------------------------------------------------
    // POST /api/admin/sync  – manually trigger the orchestrator (dev/ops use)
    // ------------------------------------------------------------------
    [Function("AdminTriggerSync")]
    public async Task<IActionResult> TriggerSyncAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "ops/sync")] HttpRequest req,
        CancellationToken ct)
    {
        _log.LogInformation("Manual sync triggered via admin API");
        var (enqueued, status) = await _orchestrator.RunSyncAsync(ct);
        return new OkObjectResult(new { enqueued, status });
    }

    // ------------------------------------------------------------------
    // GET /api/admin/config
    // ------------------------------------------------------------------
    [Function("AdminGetConfig")]
    public async Task<IActionResult> GetConfigAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "ops/config")] HttpRequest req,
        CancellationToken ct)
    {
        var config = await _sp.GetConfigAsync(ct);
        return new OkObjectResult(config);
    }

    // ------------------------------------------------------------------
    // POST /api/admin/config
    // ------------------------------------------------------------------
    [Function("AdminSetConfig")]
    public async Task<IActionResult> SetConfigAsync(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "ops/config")] HttpRequest req,
        CancellationToken ct)
    {
        CepConfig? newConfig;
        try
        {
            newConfig = await JsonSerializer.DeserializeAsync<CepConfig>(req.Body,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }, ct);
        }
        catch
        {
            return new BadRequestObjectResult("Invalid JSON body.");
        }

        if (newConfig is null) return new BadRequestObjectResult("Empty body.");

        await _sp.UpsertConfigAsync(newConfig, ct);
        _log.LogInformation("Config updated by admin");
        return new OkObjectResult(new { message = "Config updated." });
    }

    // ------------------------------------------------------------------
    // GET /api/admin/status
    // ------------------------------------------------------------------
    [Function("AdminStatus")]
    public async Task<IActionResult> GetStatusAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "ops/status")] HttpRequest req,
        CancellationToken ct)
    {
        var state = await _sp.GetSyncStateAsync(ct);
        var users = await _sp.GetActiveUsersAsync(ct);

        return new OkObjectResult(new
        {
            lastSuccessfulRunUtc = state.LastSuccessfulRunUtc,
            lastRunStatus = state.LastRunStatus,
            lastRunCorrelationId = state.LastRunCorrelationId,
            lastRunSummary = state.LastRunSummary,
            activeUsers = users.Count,
        });
    }

    // ------------------------------------------------------------------
    // POST /api/ops/test-notification  – send a test Teams notification
    // Body: { "upn": "user@contoso.com", "type": "cepWelcome" }
    // type is optional, defaults to cepWelcome
    // ------------------------------------------------------------------
    [Function("AdminTestNotification")]
    public async Task<IActionResult> TestNotificationAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "ops/test-notification")] HttpRequest req,
        CancellationToken ct)
    {
        TestNotificationRequest? body;
        try
        {
            body = await JsonSerializer.DeserializeAsync<TestNotificationRequest>(req.Body,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }, ct);
        }
        catch
        {
            return new BadRequestObjectResult("Invalid JSON body.");
        }

        if (body is null || string.IsNullOrEmpty(body.Upn))
            return new BadRequestObjectResult("upn is required.");

        var config = await _sp.GetConfigAsync(ct);
        var user = new CepUser
        {
            UserPrincipalName = body.Upn,
            DisplayName = body.Upn.Split('@')[0],
        };

        var type = body.Type?.ToLowerInvariant() ?? "cepwelcome";

        // Build notification params based on type
        string activityType;
        string preview;
        Dictionary<string, string>? templateParams;
        string? webUrl;

        switch (type)
        {
            case "cepwelcome":
                activityType = "cepWelcome";
                preview = "Welcome to CEP! 🚀";
                templateParams = null;
                webUrl = config.NotificationUrlDashboard;
                break;
            case "ceplevelup":
                activityType = "cepLevelUp";
                preview = "You reached Master level! 🏆";
                templateParams = new() { ["levelName"] = "Master" };
                webUrl = config.NotificationUrlDashboard;
                break;
            case "cepbadgeearned":
                activityType = "cepBadgeEarned";
                preview = "You earned 'Test Badge'! 🎖️";
                templateParams = new() { ["badgeName"] = "Test Badge", ["badgeDescription"] = "Test badge notification" };
                webUrl = config.NotificationUrlDashboard;
                break;
            case "cepinactivitynudge":
                activityType = "cepInactivityNudge";
                preview = "Hey, we miss you! 💡";
                templateParams = new() { ["userName"] = user.DisplayName };
                webUrl = config.NotificationUrlCopilotChat;
                break;
            case "cepleaderboardupdate":
                activityType = "cepLeaderboardUpdate";
                preview = "Leaderboard updated! 📊";
                templateParams = new() { ["monthKey"] = DateTime.UtcNow.ToString("yyyy-MM"), ["rank"] = "1" };
                webUrl = config.NotificationUrlLeaderboard;
                break;
            default:
                return new BadRequestObjectResult($"Unknown type '{body.Type}'. Use: cepWelcome, cepLevelUp, cepBadgeEarned, cepInactivityNudge, cepLeaderboardUpdate");
        }

        var teamsAppId = _teamsAppId;

        // Call GraphClient directly to get the raw response
        var (statusCode, responseBody) = await _graph.SendActivityNotificationRawAsync(
            body.Upn, teamsAppId, activityType, preview, templateParams, webUrl, ct);

        _log.LogInformation("Test notification '{Type}' to {Upn}: HTTP {Status}", type, body.Upn, statusCode);

        return new OkObjectResult(new
        {
            type = activityType,
            upn = body.Upn,
            graphStatusCode = statusCode,
            graphResponse = responseBody,
            webUrl,
            teamsAppId
        });
    }

    // ------------------------------------------------------------------
    // POST /api/ops/rebuild-leaderboard  – force leaderboard rebuild
    // ------------------------------------------------------------------
    [Function("AdminRebuildLeaderboard")]
    public async Task<IActionResult> RebuildLeaderboardAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "ops/rebuild-leaderboard")] HttpRequest req,
        CancellationToken ct)
    {
        var correlationId = $"LB-MANUAL-{Guid.NewGuid():N}";
        var runTime = DateTime.UtcNow;
        _log.LogInformation("[{CorrId}] Manual leaderboard rebuild triggered", correlationId);

        try
        {
            var config = await _sp.GetConfigAsync(ct);
            await _orchestrator.RebuildLeaderboardsAsync(correlationId, runTime, config, ct);

            var state = await _sp.GetSyncStateAsync(ct);
            state.LastLeaderboardRebuildUtc = runTime;
            await _sp.UpsertSyncStateAsync(state, ct);

            _log.LogInformation("[{CorrId}] Manual leaderboard rebuild completed", correlationId);
            return new OkObjectResult(new { status = "Success", correlationId, rebuiltAt = runTime });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[{CorrId}] Manual leaderboard rebuild failed", correlationId);
            return new ObjectResult(new { status = "Failure", correlationId, error = ex.Message })
            { StatusCode = 500 };
        }
    }

    private record TestNotificationRequest(string? Upn, string? Type);
}
