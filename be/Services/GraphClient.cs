using System.Net.Http.Headers;
using System.Text.Json;
using CepFunctions.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;

namespace CepFunctions.Services;

/// <summary>
/// Wraps Microsoft Graph calls needed by CEP.
/// Uses MSAL confidential client (client credentials) – no user context required.
/// </summary>
public class GraphClient
{
    private readonly IConfidentialClientApplication _app;
    private readonly HttpClient _http;
    private readonly ILogger<GraphClient> _log;

    private static readonly string[] GraphScopes = ["https://graph.microsoft.com/.default"];
    private const string GraphBase = "https://graph.microsoft.com/v1.0";

    public GraphClient(IConfiguration cfg, ILogger<GraphClient> log)
    {
        _log = log;
        _app = ConfidentialClientApplicationBuilder
            .Create(cfg["ClientId"])
            .WithClientSecret(cfg["ClientSecret"])
            .WithAuthority($"https://login.microsoftonline.com/{cfg["TenantId"]}")
            .Build();

        _http = new HttpClient();
        _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    // ------------------------------------------------------------------
    // AI Interaction History
    // ------------------------------------------------------------------

    /// <summary>
    /// Fetches all enterprise interactions for a user in [from, to) UTC window.
    /// Counts only userPrompt interactions, grouped by app and day.
    /// Returns aggregates: (usageDate, appKey) -> promptCount.
    /// </summary>
    public async Task<Dictionary<(DateOnly Date, string AppKey), int>> GetDailyPromptCountsAsync(
        string aadUserId,
        DateTime fromUtc,
        DateTime toUtc,
        CancellationToken ct = default)
    {
        var token = await GetTokenAsync();
        var aggregates = new Dictionary<(DateOnly, string), int>();

        var filter = $"createdDateTime gt {fromUtc:o} and createdDateTime lt {toUtc:o}";
        var url = $"{GraphBase}/copilot/users/{aadUserId}/interactionHistory/getAllEnterpriseInteractions" +
                  $"?$top=100&$filter={Uri.EscapeDataString(filter)}" +
                  $"&$select=id,createdDateTime,interactionType,appClass";

        int page = 0;
        while (!string.IsNullOrEmpty(url))
        {
            page++;
            _log.LogDebug("Graph AI interactions page {Page} for user {UserId}", page, aadUserId);

            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var resp = await SendWithRetryAsync(req, ct);
            resp.EnsureSuccessStatusCode();

            var json = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("value", out var arr))
            {
                foreach (var item in arr.EnumerateArray())
                {
                    // Privacy: only count prompts, never read body/content
                    var interactionType = item.TryGetProperty("interactionType", out var it) ? it.GetString() : null;
                    if (!string.Equals(interactionType, "userPrompt", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var appClass = item.TryGetProperty("appClass", out var ac) ? ac.GetString() ?? "" : "";
                    var created = item.TryGetProperty("createdDateTime", out var cd) && cd.TryGetDateTime(out var dt)
                        ? dt.ToUniversalTime()
                        : (DateTime?)null;

                    if (created is null) continue;

                    var key = (DateOnly.FromDateTime(created.Value), CepActivityLog.NormaliseAppClass(appClass));
                    aggregates[key] = aggregates.GetValueOrDefault(key) + 1;
                }
            }

            // Pagination
            url = root.TryGetProperty("@odata.nextLink", out var nl) ? nl.GetString() ?? "" : "";
        }

        return aggregates;
    }

    // ------------------------------------------------------------------
    // Teams Activity Notifications
    // ------------------------------------------------------------------

    /// <summary>
    /// Sends a Teams activity notification to a single user.
    /// activityType should match the Teams app manifest, or use "systemDefault".
    /// </summary>
    public async Task SendActivityNotificationAsync(
        string upn,
        string teamsAppId,
        string activityType,
        string previewText,
        CancellationToken ct = default)
    {
        var token = await GetTokenAsync();
        var url = $"{GraphBase}/users/{Uri.EscapeDataString(upn)}/teamwork/sendActivityNotification";

        var body = new
        {
            topic = new { source = "entityUrl", value = $"https://teams.microsoft.com/l/app/{teamsAppId}" },
            activityType,
            previewText = new { content = previewText },
            recipient = new { @odata_type = "#microsoft.graph.aadUserNotificationRecipient", userId = upn }
        };

        var payload = JsonSerializer.Serialize(body, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });

        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        req.Content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");

        var resp = await SendWithRetryAsync(req, ct);
        if (!resp.IsSuccessStatusCode)
            _log.LogWarning("Teams notification failed for {Upn}: {Status}", upn, resp.StatusCode);
    }

    // ------------------------------------------------------------------
    // Internals
    // ------------------------------------------------------------------

    private async Task<string> GetTokenAsync()
    {
        var result = await _app.AcquireTokenForClient(GraphScopes).ExecuteAsync();
        return result.AccessToken;
    }

    /// <summary>Sends a request and retries once on 429 (Too Many Requests), honouring Retry-After.</summary>
    private async Task<HttpResponseMessage> SendWithRetryAsync(HttpRequestMessage req, CancellationToken ct)
    {
        var resp = await _http.SendAsync(req, ct);

        if ((int)resp.StatusCode == 429)
        {
            var retryAfter = resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(10);
            _log.LogWarning("Graph throttled (429). Waiting {Seconds}s before retry.", retryAfter.TotalSeconds);
            await Task.Delay(retryAfter, ct);

            // Re-create the request (HttpRequestMessage cannot be reused)
            using var retry = new HttpRequestMessage(req.Method, req.RequestUri);
            foreach (var h in req.Headers) retry.Headers.TryAddWithoutValidation(h.Key, h.Value);
            retry.Content = req.Content;
            resp = await _http.SendAsync(retry, ct);
        }

        return resp;
    }
}
