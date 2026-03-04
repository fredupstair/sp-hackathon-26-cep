namespace CepFunctions.Models;

/// <summary>
/// Materialised monthly ranking row in CEP_Leaderboard.
/// One row = one user + one month + one scope (Global / Department / Team).
/// </summary>
public class CepLeaderboardEntry
{
    public string? SpItemId { get; set; }

    /// <summary>YYYY-MM</summary>
    public string MonthKey { get; set; } = "";

    /// <summary>Global | Department:&lt;name&gt; | Team:&lt;name&gt;</summary>
    public string Scope { get; set; } = "Global";

    public string AadUserId { get; set; } = "";
    public string UserEmail { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string Department { get; set; } = "";
    public string Team { get; set; } = "";
    public int MonthlyPoints { get; set; }
    public int Rank { get; set; }
    public string Level { get; set; } = "Bronze";

    public Dictionary<string, object?> ToSpFields() => new()
    {
        ["Title"] = MonthKey,
        ["MonthKey"] = MonthKey,
        ["Scope"] = Scope,
        ["AadUserId"] = AadUserId,
        ["UserEmail"] = UserEmail,
        ["DisplayName"] = DisplayName,
        ["Department"] = Department,
        ["Team"] = Team,
        ["MonthlyPoints"] = MonthlyPoints,
        ["Rank"] = Rank,
        ["Level"] = Level,
    };

    public static CepLeaderboardEntry FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        MonthKey = f.Str("MonthKey"),
        Scope = f.Str("Scope", "Global"),
        AadUserId = f.Str("AadUserId"),
        UserEmail = f.Str("UserEmail"),
        DisplayName = f.Str("DisplayName"),
        Department = f.Str("Department"),
        Team = f.Str("Team"),
        MonthlyPoints = f.Int("MonthlyPoints"),
        Rank = f.Int("Rank"),
        Level = f.Str("Level", "Bronze"),
    };
}
