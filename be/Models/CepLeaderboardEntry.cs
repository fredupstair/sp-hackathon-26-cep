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
        ["CEP_LB_MonthKey"] = MonthKey,
        ["CEP_LB_Scope"] = Scope,
        ["CEP_LB_AadUserId"] = AadUserId,
        ["CEP_LB_UserEmail"] = UserEmail,
        ["CEP_LB_DisplayName"] = DisplayName,
        ["CEP_LB_Department"] = Department,
        ["CEP_LB_Team"] = Team,
        ["CEP_LB_MonthlyPoints"] = MonthlyPoints,
        ["CEP_LB_Rank"] = Rank,
        ["CEP_LB_Level"] = Level,
    };

    public static CepLeaderboardEntry FromSpFields(string spItemId, Dictionary<string, object?> f) => new()
    {
        SpItemId = spItemId,
        MonthKey = f.Str("CEP_LB_MonthKey"),
        Scope = f.Str("CEP_LB_Scope", "Global"),
        AadUserId = f.Str("CEP_LB_AadUserId"),
        UserEmail = f.Str("CEP_LB_UserEmail"),
        DisplayName = f.Str("CEP_LB_DisplayName"),
        Department = f.Str("CEP_LB_Department"),
        Team = f.Str("CEP_LB_Team"),
        MonthlyPoints = f.Int("CEP_LB_MonthlyPoints"),
        Rank = f.Int("CEP_LB_Rank"),
        Level = f.Str("CEP_LB_Level", "Bronze"),
    };
}
