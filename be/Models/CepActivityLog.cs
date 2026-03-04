namespace CepFunctions.Models;

/// <summary>
/// Daily usage aggregate stored in CEP_ActivityLog.
/// One row = one user + one day + one app.
/// </summary>
public class CepActivityLog
{
    public string? SpItemId { get; set; }

    public string AadUserId { get; set; } = "";
    public string UserEmail { get; set; } = "";

    /// <summary>Date-only reference day (UTC).</summary>
    public DateTime UsageDate { get; set; }

    /// <summary>Normalised app key: Word / Excel / PowerPoint / Outlook / Teams / OneNote / Loop / M365Chat / Other.</summary>
    public string AppKey { get; set; } = "";

    public int PromptCount { get; set; }
    public int PointsEarned { get; set; }

    /// <summary>YYYY-MM – used for efficient monthly queries.</summary>
    public string MonthKey { get; set; } = "";

    public Dictionary<string, object?> ToSpFields() => new()
    {
        ["Title"] = $"{AadUserId}_{MonthKey}_{AppKey}",
        ["AadUserId"] = AadUserId,
        ["UserEmail"] = UserEmail,
        ["UsageDate"] = UsageDate.ToString("yyyy-MM-dd"),
        ["AppKey"] = AppKey,
        ["PromptCount"] = PromptCount,
        ["PointsEarned"] = PointsEarned,
        ["MonthKey"] = MonthKey,
    };

    public static CepActivityLog FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        AadUserId = f.Str("AadUserId"),
        UserEmail = f.Str("UserEmail"),
        UsageDate = f.Dt("UsageDate") ?? DateTime.UtcNow.Date,
        AppKey = f.Str("AppKey"),
        PromptCount = f.Int("PromptCount"),
        PointsEarned = f.Int("PointsEarned"),
        MonthKey = f.Str("MonthKey"),
    };

    /// <summary>
    /// Maps a Graph appClass string to the canonical AppKey.
    /// </summary>
    public static string NormaliseAppClass(string appClass) =>
        appClass switch
        {
            _ when appClass.Contains(".Copilot.Word", StringComparison.OrdinalIgnoreCase) => "Word",
            _ when appClass.Contains(".Copilot.Excel", StringComparison.OrdinalIgnoreCase) => "Excel",
            _ when appClass.Contains(".Copilot.PowerPoint", StringComparison.OrdinalIgnoreCase) => "PowerPoint",
            _ when appClass.Contains(".Copilot.Outlook", StringComparison.OrdinalIgnoreCase) => "Outlook",
            _ when appClass.Contains(".Copilot.Teams", StringComparison.OrdinalIgnoreCase) => "Teams",
            _ when appClass.Contains(".Copilot.OneNote", StringComparison.OrdinalIgnoreCase) => "OneNote",
            _ when appClass.Contains(".Copilot.Loop", StringComparison.OrdinalIgnoreCase) => "Loop",
            _ when appClass.Contains(".Copilot.BizChat", StringComparison.OrdinalIgnoreCase) => "M365Chat",
            _ => "Other",
        };
}
