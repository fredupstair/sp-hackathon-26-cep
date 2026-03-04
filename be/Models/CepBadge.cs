namespace CepFunctions.Models;

public class CepBadge
{
    public string? SpItemId { get; set; }

    public string AadUserId { get; set; } = "";
    public string UserEmail { get; set; } = "";

    /// <summary>Stable technical key – used for duplicate detection.</summary>
    public string BadgeKey { get; set; } = "";

    public string BadgeName { get; set; } = "";
    public string Description { get; set; } = "";
    public DateTime? EarnedDate { get; set; }
    public string MonthKey { get; set; } = "";

    public Dictionary<string, object?> ToSpFields() => new()
    {
        ["Title"] = BadgeName,
        ["CEP_Badge_AadUserId"] = AadUserId,
        ["CEP_Badge_UserEmail"] = UserEmail,
        ["CEP_Badge_BadgeKey"] = BadgeKey,
        ["CEP_Badge_BadgeType"] = BadgeName,
        ["CEP_Badge_Description"] = Description,
        ["CEP_Badge_EarnedDate"] = EarnedDate?.ToUniversalTime().ToString("o"),
        ["CEP_Badge_MonthKey"] = MonthKey,
    };

    public static CepBadge FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        AadUserId = f.Str("CEP_Badge_AadUserId"),
        UserEmail = f.Str("CEP_Badge_UserEmail"),
        BadgeKey = f.Str("CEP_Badge_BadgeKey"),
        BadgeName = f.Str("Title"),
        Description = f.Str("CEP_Badge_Description"),
        EarnedDate = f.Dt("CEP_Badge_EarnedDate"),
        MonthKey = f.Str("CEP_Badge_MonthKey"),
    };

    // Canonical badge keys
    public static class Keys
    {
        public const string FirstSteps = "FirstSteps";
        public const string CrossAppExplorer = "CrossAppExplorer";
        public const string WeeklyWarrior = "WeeklyWarrior";
        public const string MonthlyMaster = "MonthlyMaster";
        public const string ConsistencyKing = "ConsistencyKing";
    }
}
