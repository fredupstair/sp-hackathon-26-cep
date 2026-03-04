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
        ["CEP_Sync_LastSuccessfulRunUtc"] = LastSuccessfulRunUtc?.ToUniversalTime().ToString("o"),
        ["CEP_Sync_LastRunStatus"] = LastRunStatus,
        ["CEP_Sync_LastRunCorrelationId"] = LastRunCorrelationId,
        ["CEP_Sync_LastRunSummary"] = LastRunSummary,
    };

    public static CepSyncState FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        LastSuccessfulRunUtc = f.Dt("CEP_Sync_LastSuccessfulRunUtc"),
        LastRunStatus = f.Str("CEP_Sync_LastRunStatus", "Unknown"),
        LastRunCorrelationId = f.Str("CEP_Sync_LastRunCorrelationId"),
        LastRunSummary = f.Str("CEP_Sync_LastRunSummary"),
    };
}
