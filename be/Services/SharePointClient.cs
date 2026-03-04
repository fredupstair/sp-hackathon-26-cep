using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using CepFunctions.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;

namespace CepFunctions.Services;

/// <summary>
/// Thin client for SharePoint Lists via Microsoft Graph REST (/sites/{id}/lists/{id}/items).
/// Uses the same app credential as GraphClient.
/// All list item payloads use the "fields" wrapper pattern.
/// </summary>
public class SharePointClient
{
    private readonly IConfidentialClientApplication _app;
    private readonly HttpClient _http;
    private readonly ILogger<SharePointClient> _log;
    private readonly string _siteId;

    // List IDs injected from configuration
    private readonly string _listUsers;
    private readonly string _listActivityLog;
    private readonly string _listLeaderboard;
    private readonly string _listBadges;
    private readonly string _listConfig;
    private readonly string _listSyncState;

    private static readonly string[] GraphScopes = ["https://graph.microsoft.com/.default"];
    private const string GraphBase = "https://graph.microsoft.com/v1.0";

    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    public SharePointClient(IConfiguration cfg, ILogger<SharePointClient> log)
    {
        _log = log;
        _siteId = cfg["SpSiteId"] ?? throw new InvalidOperationException("SpSiteId missing");
        _listUsers = cfg["ListId_Users"] ?? throw new InvalidOperationException("ListId_Users missing");
        _listActivityLog = cfg["ListId_ActivityLog"] ?? throw new InvalidOperationException("ListId_ActivityLog missing");
        _listLeaderboard = cfg["ListId_Leaderboard"] ?? throw new InvalidOperationException("ListId_Leaderboard missing");
        _listBadges = cfg["ListId_Badges"] ?? throw new InvalidOperationException("ListId_Badges missing");
        _listConfig = cfg["ListId_Config"] ?? throw new InvalidOperationException("ListId_Config missing");
        _listSyncState = cfg["ListId_SyncState"] ?? throw new InvalidOperationException("ListId_SyncState missing");

        _app = ConfidentialClientApplicationBuilder
            .Create(cfg["ClientId"])
            .WithClientSecret(cfg["ClientSecret"])
            .WithAuthority($"https://login.microsoftonline.com/{cfg["TenantId"]}")
            .Build();

        _http = new HttpClient();
    }

    // ------------------------------------------------------------------
    // CEP_Config
    // ------------------------------------------------------------------

    public async Task<CepConfig> GetConfigAsync(CancellationToken ct = default)
    {
        var items = await GetAllFieldsAsync(_listConfig, ct: ct);
        return CepConfig.FromSpItems(items);
    }

    public async Task UpsertConfigAsync(CepConfig config, CancellationToken ct = default)
    {
        // Config is a single row now - upsert the whole row
        var existing = await GetAllItemsAsync(_listConfig, ct: ct);
        var first = existing.FirstOrDefault();
        foreach (var row in config.ToSpRows())
        {
            if (first is not null && first.TryGetValue("id", out var spid))
                await PatchItemAsync(_listConfig, spid?.ToString() ?? "", row, ct);
            else
                await CreateItemAsync(_listConfig, row, ct);
        }
    }

    // ------------------------------------------------------------------
    // CEP_SyncState
    // ------------------------------------------------------------------

    public async Task<CepSyncState> GetSyncStateAsync(CancellationToken ct = default)
    {
        var items = await GetAllItemsAsync(_listSyncState, ct: ct);
        var first = items.FirstOrDefault();
        if (first is null) return new CepSyncState();
        var id = first.TryGetValue("id", out var v) ? v?.ToString() ?? "" : "";
        return CepSyncState.FromSpFields(id, first);
    }

    public async Task UpsertSyncStateAsync(CepSyncState state, CancellationToken ct = default)
    {
        var fields = state.ToSpFields();
        if (!string.IsNullOrEmpty(state.SpItemId))
            await PatchItemAsync(_listSyncState, state.SpItemId, fields, ct);
        else
            await CreateItemAsync(_listSyncState, fields, ct);
    }

    // ------------------------------------------------------------------
    // CEP_Users
    // ------------------------------------------------------------------

    public async Task<List<CepUser>> GetActiveUsersAsync(CancellationToken ct = default)
    {
        // Note: filtering CEP_IsActive (Boolean, non-indexed) via OData on SharePoint Graph
        // silently returns empty results even with the HonorNonIndexed Prefer header.
        // Fetch all users and filter in memory instead.
        var items = await GetAllItemsAsync(_listUsers, ct: ct);
        return items.Select(i =>
        {
            var id = i.TryGetValue("id", out var v) ? v?.ToString() ?? "" : "";
            var fields = i.TryGetValue("fields", out var f) && f is JsonElement je
                ? JsonSerializer.Deserialize<Dictionary<string, object?>>(je.GetRawText(), JsonOpts) ?? []
                : i;
            return CepUser.FromSpFields(id, fields);
        })
        .Where(u => u.IsActive)
        .ToList();
    }

    public async Task<CepUser?> GetUserByAadIdAsync(string aadUserId, CancellationToken ct = default)
    {
        var filter = $"fields/CEP_AadUserId eq '{aadUserId}'";
        var items = await GetAllItemsAsync(_listUsers, filter, ct);
        var first = items.FirstOrDefault();
        if (first is null) return null;
        var id = first.TryGetValue("id", out var v) ? v?.ToString() ?? "" : "";
        return CepUser.FromSpFields(id, ExtractFields(first));
    }

    public async Task<CepUser?> GetUserByUpnAsync(string upn, CancellationToken ct = default)
    {
        var filter = $"fields/CEP_UserPrincipalName eq '{upn}'";
        var items = await GetAllItemsAsync(_listUsers, filter, ct);
        var first = items.FirstOrDefault();
        if (first is null) return null;
        var id = first.TryGetValue("id", out var v) ? v?.ToString() ?? "" : "";
        return CepUser.FromSpFields(id, ExtractFields(first));
    }

    public async Task<CepUser> UpsertUserAsync(CepUser user, CancellationToken ct = default)
    {
        var fields = user.ToSpFields();
        if (!string.IsNullOrEmpty(user.SpItemId))
        {
            await PatchItemAsync(_listUsers, user.SpItemId, fields, ct);
            return user;
        }
        var created = await CreateItemAsync(_listUsers, fields, ct);
        user.SpItemId = created;
        return user;
    }

    // ------------------------------------------------------------------
    // CEP_ActivityLog
    // ------------------------------------------------------------------

    /// <summary>
    /// Returns existing aggregates for a user in the given month.
    /// Key = (AppKey, UsageDate.Date).
    /// </summary>
    public async Task<Dictionary<(string AppKey, DateOnly Date), CepActivityLog>> GetActivityLogsForUserMonthAsync(
        string aadUserId, string monthKey, CancellationToken ct = default)
    {
        var filter = $"fields/CEP_Log_AadUserId eq '{aadUserId}' and fields/CEP_Log_MonthKey eq '{monthKey}'";
        var items = await GetAllItemsAsync(_listActivityLog, filter, ct);
        return items
            .Select(i =>
            {
                var id = i.TryGetValue("id", out var v) ? v?.ToString() ?? "" : "";
                return CepActivityLog.FromSpFields(id, ExtractFields(i));
            })
            .ToDictionary(l => (l.AppKey, l.UsageDate != default ? DateOnly.FromDateTime(l.UsageDate) : DateOnly.MinValue));
    }

    public async Task UpsertActivityLogAsync(CepActivityLog log, CancellationToken ct = default)
    {
        var fields = log.ToSpFields();
        if (!string.IsNullOrEmpty(log.SpItemId))
            await PatchItemAsync(_listActivityLog, log.SpItemId, fields, ct);
        else
            await CreateItemAsync(_listActivityLog, fields, ct);
    }

    // ------------------------------------------------------------------
    // CEP_Leaderboard
    // ------------------------------------------------------------------

    /// <summary>Deletes all existing rows for a given month+scope, then bulk-creates new ones.</summary>
    public async Task ReplaceLeaderboardAsync(string monthKey, string scope, IEnumerable<CepLeaderboardEntry> entries, CancellationToken ct = default)
    {
        var filter = $"fields/CEP_LB_MonthKey eq '{monthKey}' and fields/CEP_LB_Scope eq '{scope}'";
        var existing = await GetAllItemsAsync(_listLeaderboard, filter, ct);
        foreach (var item in existing)
        {
            if (item.TryGetValue("id", out var v) && v?.ToString() is { } id)
                await DeleteItemAsync(_listLeaderboard, id, ct);
        }
        foreach (var entry in entries)
            await CreateItemAsync(_listLeaderboard, entry.ToSpFields(), ct);
    }

    public async Task<List<CepLeaderboardEntry>> GetLeaderboardAsync(
        string monthKey, string scope, int page = 1, int pageSize = 20, CancellationToken ct = default)
    {
        var filter = $"fields/CEP_LB_MonthKey eq '{monthKey}' and fields/CEP_LB_Scope eq '{scope}'";
        var items = await GetAllItemsAsync(_listLeaderboard, filter, ct);
        return items
            .Select(i => CepLeaderboardEntry.FromSpFields(
                i.TryGetValue("id", out var v) ? v?.ToString() ?? "" : "",
                ExtractFields(i)))
            .OrderBy(e => e.Rank)
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .ToList();
    }

    // ------------------------------------------------------------------
    // CEP_Badges
    // ------------------------------------------------------------------

    public async Task<List<CepBadge>> GetUserBadgesAsync(string aadUserId, CancellationToken ct = default)
    {
        var filter = $"fields/CEP_Badge_AadUserId eq '{aadUserId}'";
        var items = await GetAllItemsAsync(_listBadges, filter, ct);
        return items.Select(i => CepBadge.FromSpFields(
            i.TryGetValue("id", out var v) ? v?.ToString() ?? "" : "",
            ExtractFields(i))).ToList();
    }

    /// <summary>Creates a badge only if it doesn't already exist (idempotent).</summary>
    public async Task<bool> TryAwardBadgeAsync(CepBadge badge, CancellationToken ct = default)
    {
        var filter = $"fields/CEP_Badge_AadUserId eq '{badge.AadUserId}' and fields/CEP_Badge_BadgeKey eq '{badge.BadgeKey}'";
        if (!string.IsNullOrEmpty(badge.MonthKey))
            filter += $" and fields/CEP_Badge_MonthKey eq '{badge.MonthKey}'";

        var existing = await GetAllItemsAsync(_listBadges, filter, ct);
        if (existing.Count > 0) return false; // Already awarded

        await CreateItemAsync(_listBadges, badge.ToSpFields(), ct);
        return true;
    }

    // ------------------------------------------------------------------
    // Low-level Graph SP REST helpers
    // ------------------------------------------------------------------

    private async Task<List<Dictionary<string, object?>>> GetAllItemsAsync(
        string listId, string? filter = null, CancellationToken ct = default)
    {
        var token = await GetTokenAsync();
        var results = new List<Dictionary<string, object?>>();

        var url = $"{GraphBase}/sites/{_siteId}/lists/{listId}/items?expand=fields&$top=999";
        if (!string.IsNullOrEmpty(filter)) url += $"&$filter={Uri.EscapeDataString(filter)}";

        while (!string.IsNullOrEmpty(url))
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            // Allow non-indexed column filters (SP lists < 5000 rows in this hackathon context)
            req.Headers.TryAddWithoutValidation("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var resp = await _http.SendAsync(req, ct);
            resp.EnsureSuccessStatusCode();

            using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct));
            var root = doc.RootElement;

            if (root.TryGetProperty("value", out var arr))
                foreach (var item in arr.EnumerateArray())
                    results.Add(JsonSerializer.Deserialize<Dictionary<string, object?>>(item.GetRawText(), JsonOpts) ?? []);

            url = root.TryGetProperty("@odata.nextLink", out var nl) ? nl.GetString() ?? "" : "";
        }

        return results;
    }

    private async Task<List<Dictionary<string, object?>>> GetAllFieldsAsync(
        string listId, string? filter = null, CancellationToken ct = default)
    {
        var items = await GetAllItemsAsync(listId, filter, ct);
        return items.Select(i =>
        {
            var fields = ExtractFields(i);
            if (i.TryGetValue("id", out var id)) fields["id"] = id;
            return fields;
        }).ToList();
    }

    private async Task<string> CreateItemAsync(
        string listId, Dictionary<string, object?> fields, CancellationToken ct)
    {
        var token = await GetTokenAsync();
        var url = $"{GraphBase}/sites/{_siteId}/lists/{listId}/items";
        var body = JsonSerializer.Serialize(new { fields });

        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Content = new StringContent(body, Encoding.UTF8, "application/json");

        var resp = await _http.SendAsync(req, ct);
        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException(
                $"Graph POST {url} → {(int)resp.StatusCode}: {errBody}");
        }

        using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct));
        return doc.RootElement.TryGetProperty("id", out var id) ? id.GetString() ?? "" : "";
    }

    private async Task PatchItemAsync(
        string listId, string itemId, Dictionary<string, object?> fields, CancellationToken ct)
    {
        var token = await GetTokenAsync();
        var url = $"{GraphBase}/sites/{_siteId}/lists/{listId}/items/{itemId}/fields";
        var body = JsonSerializer.Serialize(fields);

        using var req = new HttpRequestMessage(HttpMethod.Patch, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Content = new StringContent(body, Encoding.UTF8, "application/json");

        var resp = await _http.SendAsync(req, ct);
        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync(ct);
            throw new HttpRequestException(
                $"Graph PATCH {url} → {(int)resp.StatusCode}: {errBody}");
        }
    }

    private async Task DeleteItemAsync(string listId, string itemId, CancellationToken ct)
    {
        var token = await GetTokenAsync();
        var url = $"{GraphBase}/sites/{_siteId}/lists/{listId}/items/{itemId}";
        using var req = new HttpRequestMessage(HttpMethod.Delete, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        var resp = await _http.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();
    }

    private async Task<string> GetTokenAsync()
    {
        var result = await _app.AcquireTokenForClient(GraphScopes).ExecuteAsync();
        return result.AccessToken;
    }

    private static Dictionary<string, object?> ExtractFields(Dictionary<string, object?> item)
    {
        if (item.TryGetValue("fields", out var f) && f is JsonElement je)
            return JsonSerializer.Deserialize<Dictionary<string, object?>>(je.GetRawText(), JsonOpts) ?? [];
        return item;
    }
}
