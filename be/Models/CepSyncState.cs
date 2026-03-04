namespace CepFunctions.Models;

public class CepSyncState
{
    public string? SpItemId { get; set; }

    public DateTime? LastSuccessfulRunUtc { get; set; }
    public string LastRunStatus { get; set; } = "Unknown"; // Success / Failure / Running
    public string LastRunCorrelationId { get; set; } = "";
    public string LastRunSummary { get; set; } = "";

    public Dictionary<string, object?> ToSpFields() => new()
    {
        ["Title"] = "SyncState",
        ["LastSuccessfulRunUtc"] = LastSuccessfulRunUtc?.ToUniversalTime().ToString("o"),
        ["LastRunStatus"] = LastRunStatus,
        ["LastRunCorrelationId"] = LastRunCorrelationId,
        ["LastRunSummary"] = LastRunSummary,
    };

    public static CepSyncState FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        LastSuccessfulRunUtc = f.Dt("LastSuccessfulRunUtc"),
        LastRunStatus = f.Str("LastRunStatus", "Unknown"),
        LastRunCorrelationId = f.Str("LastRunCorrelationId"),
        LastRunSummary = f.Str("LastRunSummary"),
    };
}
