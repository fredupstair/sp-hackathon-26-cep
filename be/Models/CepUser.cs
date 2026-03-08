using System.Text.Json.Serialization;

namespace CepFunctions.Models;

/// <summary>
/// Represents a row in the CEP_Users SharePoint list.
/// </summary>
public class CepUser
{
    [JsonPropertyName("id")]
    public string? SpItemId { get; set; }

    // Identity
    public string AadUserId { get; set; } = "";
    public string UserPrincipalName { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string Email { get; set; } = "";
    public string Department { get; set; } = "";
    public string Team { get; set; } = "";

    // Program state
    public bool IsActive { get; set; } = true;
    public DateTime? EnrollmentDate { get; set; }
    public string CurrentLevel { get; set; } = "Explorer"; // Explorer / Practitioner / Master
    public int TotalPoints { get; set; }
    public int MonthlyPoints { get; set; }

    // Activity & sync
    public DateTime? LastActivityDate { get; set; }
    public DateTime? LastSyncWatermarkUtc { get; set; }
    public DateTime? LastNudgeSentUtc { get; set; }
    public string LastLeaderboardNotifiedMonth { get; set; } = "";

    // Preferences
    public bool IsEngagementNudgesEnabled { get; set; } = true;

    /// <summary>
    /// Builds the Graph/REST field payload for a SharePoint list item create/update.
    /// </summary>
    public Dictionary<string, object?> ToSpFields() => new()
    {
        ["Title"] = DisplayName,
        ["CEP_AadUserId"] = AadUserId,
        ["CEP_UserPrincipalName"] = UserPrincipalName,
        ["CEP_UserEmail"] = Email,
        ["CEP_Department"] = Department,
        ["CEP_Team"] = Team,
        ["CEP_IsActive"] = IsActive,
        ["CEP_EnrollmentDate"] = EnrollmentDate?.ToUniversalTime().ToString("o"),
        ["CEP_CurrentLevel"] = CurrentLevel,
        ["CEP_TotalPoints"] = TotalPoints,
        ["CEP_MonthlyPoints"] = MonthlyPoints,
        ["CEP_LastActivityDate"] = LastActivityDate?.ToUniversalTime().ToString("o"),
        ["CEP_LastSyncWatermarkUtc"] = LastSyncWatermarkUtc?.ToUniversalTime().ToString("o"),
        ["CEP_LastNudgeSentUtc"] = LastNudgeSentUtc?.ToUniversalTime().ToString("o"),
        ["CEP_LastLeaderboardNotifiedMonth"] = LastLeaderboardNotifiedMonth,
        ["CEP_IsEngagementNudgesEnabled"] = IsEngagementNudgesEnabled,
    };

    public static CepUser FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        AadUserId = f.Str("CEP_AadUserId"),
        UserPrincipalName = f.Str("CEP_UserPrincipalName"),
        DisplayName = f.Str("Title"),
        Email = f.Str("CEP_UserEmail"),
        Department = f.Str("CEP_Department"),
        Team = f.Str("CEP_Team"),
        IsActive = f.Bool("CEP_IsActive", true),
        EnrollmentDate = f.Dt("CEP_EnrollmentDate"),
        CurrentLevel = f.Str("CEP_CurrentLevel", "Explorer"),
        TotalPoints = f.Int("CEP_TotalPoints"),
        MonthlyPoints = f.Int("CEP_MonthlyPoints"),
        LastActivityDate = f.Dt("CEP_LastActivityDate"),
        LastSyncWatermarkUtc = f.Dt("CEP_LastSyncWatermarkUtc"),
        LastNudgeSentUtc = f.Dt("CEP_LastNudgeSentUtc"),
        LastLeaderboardNotifiedMonth = f.Str("CEP_LastLeaderboardNotifiedMonth"),
        IsEngagementNudgesEnabled = f.Bool("CEP_IsEngagementNudgesEnabled", true),
    };
}
