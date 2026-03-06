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

    /// <summary>True for manual "Copilot Win" entries (AppKey = "Win").</summary>
    public bool IsWin { get; set; } = false;

    /// <summary>User note recorded with a Win (stored on the daily aggregate row).</summary>
    public string WinNote { get; set; } = "";

    /// <summary>Whether the user opted to share this win anonymously as a community tip.</summary>
    public bool IsShared { get; set; } = false;

    public Dictionary<string, object?> ToSpFields() => new()
    {
        ["Title"] = $"{AadUserId}_{MonthKey}_{AppKey}",
        ["CEP_Log_AadUserId"] = AadUserId,
        ["CEP_Log_UserEmail"] = UserEmail,
        ["CEP_Log_UsageDate"] = UsageDate.ToString("yyyy-MM-dd"),
        ["CEP_Log_AppKey"] = AppKey,
        ["CEP_Log_PromptCount"] = PromptCount,
        ["CEP_Log_PointsEarned"] = PointsEarned,
        ["CEP_Log_MonthKey"] = MonthKey,
        ["CEP_Log_IsWin"] = IsWin,
        ["CEP_Log_WinNote"] = WinNote,
        ["CEP_Log_IsShared"] = IsShared,
    };

    public static CepActivityLog FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        AadUserId = f.Str("CEP_Log_AadUserId"),
        UserEmail = f.Str("CEP_Log_UserEmail"),
        UsageDate = f.Dt("CEP_Log_UsageDate") ?? DateTime.UtcNow.Date,
        AppKey = f.Str("CEP_Log_AppKey"),
        PromptCount = f.Int("CEP_Log_PromptCount"),
        PointsEarned = f.Int("CEP_Log_PointsEarned"),
        MonthKey = f.Str("CEP_Log_MonthKey"),
        IsWin = f.Bool("CEP_Log_IsWin"),
        WinNote = f.Str("CEP_Log_WinNote"),
        IsShared = f.Bool("CEP_Log_IsShared"),
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
