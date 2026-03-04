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
    public string CurrentLevel { get; set; } = "Bronze"; // Bronze / Silver / Gold
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
        ["AadUserId"] = AadUserId,
        ["UserPrincipalName"] = UserPrincipalName,
        ["UserEmail"] = Email,
        ["Department"] = Department,
        ["Team"] = Team,
        ["IsActive"] = IsActive,
        ["EnrollmentDate"] = EnrollmentDate?.ToUniversalTime().ToString("o"),
        ["CurrentLevel"] = CurrentLevel,
        ["TotalPoints"] = TotalPoints,
        ["MonthlyPoints"] = MonthlyPoints,
        ["LastActivityDate"] = LastActivityDate?.ToUniversalTime().ToString("o"),
        ["LastSyncWatermarkUtc"] = LastSyncWatermarkUtc?.ToUniversalTime().ToString("o"),
        ["LastNudgeSentUtc"] = LastNudgeSentUtc?.ToUniversalTime().ToString("o"),
        ["LastLeaderboardNotifiedMonth"] = LastLeaderboardNotifiedMonth,
        ["IsEngagementNudgesEnabled"] = IsEngagementNudgesEnabled,
    };

    public static CepUser FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        AadUserId = f.Str("AadUserId"),
        UserPrincipalName = f.Str("UserPrincipalName"),
        DisplayName = f.Str("Title"),
        Email = f.Str("UserEmail"),
        Department = f.Str("Department"),
        Team = f.Str("Team"),
        IsActive = f.Bool("IsActive", true),
        EnrollmentDate = f.Dt("EnrollmentDate"),
        CurrentLevel = f.Str("CurrentLevel", "Bronze"),
        TotalPoints = f.Int("TotalPoints"),
        MonthlyPoints = f.Int("MonthlyPoints"),
        LastActivityDate = f.Dt("LastActivityDate"),
        LastSyncWatermarkUtc = f.Dt("LastSyncWatermarkUtc"),
        LastNudgeSentUtc = f.Dt("LastNudgeSentUtc"),
        LastLeaderboardNotifiedMonth = f.Str("LastLeaderboardNotifiedMonth"),
        IsEngagementNudgesEnabled = f.Bool("IsEngagementNudgesEnabled", true),
    };
}
