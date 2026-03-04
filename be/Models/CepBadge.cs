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
        ["AadUserId"] = AadUserId,
        ["UserEmail"] = UserEmail,
        ["BadgeKey"] = BadgeKey,
        ["Description"] = Description,
        ["EarnedDate"] = EarnedDate?.ToUniversalTime().ToString("o"),
        ["MonthKey"] = MonthKey,
    };

    public static CepBadge FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        AadUserId = f.Str("AadUserId"),
        UserEmail = f.Str("UserEmail"),
        BadgeKey = f.Str("BadgeKey"),
        BadgeName = f.Str("Title"),
        Description = f.Str("Description"),
        EarnedDate = f.Dt("EarnedDate"),
        MonthKey = f.Str("MonthKey"),
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
