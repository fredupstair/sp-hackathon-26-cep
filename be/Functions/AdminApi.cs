using System.Text.Json;
using CepFunctions.Models;
using CepFunctions.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
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
    private readonly ILogger<AdminApi> _log;

    public AdminApi(SharePointClient sp, OrchestratorTimer orchestrator, ILogger<AdminApi> log)
    {
        _sp = sp;
        _orchestrator = orchestrator;
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
}
