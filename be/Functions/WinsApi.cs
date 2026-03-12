using CepFunctions.Models;
using CepFunctions.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// Wins API – public feed of shared Copilot Wins from all enrolled users.
///
/// GET /api/wins/shared?month=YYYY-MM[&amp;appKey=Teams]
/// </summary>
public class WinsApi
{
    private readonly SharePointClient _sp;
    private readonly ILogger<WinsApi> _log;

    public WinsApi(SharePointClient sp, ILogger<WinsApi> log)
    {
        _sp = sp;
        _log = log;
    }

    // ------------------------------------------------------------------
    // GET /api/wins/shared?month=YYYY-MM[&appKey=Teams]
    // ------------------------------------------------------------------

    [Function("GetSharedWins")]
    public async Task<IActionResult> GetSharedWinsAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "wins/shared")] HttpRequest req,
        CancellationToken ct)
    {
        var month = req.Query["month"].FirstOrDefault() ?? DateTime.UtcNow.ToString("yyyy-MM");
        var appKey = req.Query["appKey"].FirstOrDefault();
        var callerId = GetCallerId(req);   // exclude caller's own shared wins from community feed

        var wins = await _sp.GetSharedWinsAsync(month, appKey, callerId, ct);

        // Batch-load display names (one SP lookup per unique user)
        var displayNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var uniqueUserIds = wins.Select(w => w.AadUserId).Distinct().ToList();
        foreach (var uid in uniqueUserIds)
        {
            var user = await _sp.GetUserByAadIdAsync(uid, ct);
            if (user is not null)
            {
                // Privacy: keep "First L." format only
                var parts = user.DisplayName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                displayNames[uid] = parts.Length > 1
                    ? $"{parts[0]} {parts[^1][0]}."
                    : parts[0];
            }
        }

        return new OkObjectResult(new
        {
            month,
            items = wins.Select(w => new
            {
                id = w.SpItemId,
                date = w.UsageDate.ToString("yyyy-MM-dd"),
                appKey = w.WinAppKey,
                note = w.WinNote,
                displayName = displayNames.TryGetValue(w.AadUserId, out var dn) ? dn : "Un collega",
            })
        });
    }

    private static string? GetCallerId(HttpRequest req)
    {
        if (req.Headers.TryGetValue("X-MS-CLIENT-PRINCIPAL-ID", out var easyAuth) &&
            easyAuth.FirstOrDefault() is { Length: > 0 } oid)
            return oid;
        return req.Headers.TryGetValue("X-User-Id", out var devId) ? devId.FirstOrDefault() : null;
    }
}
