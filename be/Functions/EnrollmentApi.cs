using System.Text.Json;
using CepFunctions.Models;
using CepFunctions.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace CepFunctions.Functions;

/// <summary>
/// Enrollment HTTP API.
/// POST /api/enrollment/join   – opt-in
/// POST /api/enrollment/leave  – soft opt-out
/// </summary>
public class EnrollmentApi
{
    private readonly SharePointClient _sp;
    private readonly TeamsNotifier _notifier;
    private readonly ILogger<EnrollmentApi> _log;

    public EnrollmentApi(SharePointClient sp, TeamsNotifier notifier, ILogger<EnrollmentApi> log)
    {
        _sp = sp;
        _notifier = notifier;
        _log = log;
    }

    // ------------------------------------------------------------------
    // POST /api/enrollment/join
    // ------------------------------------------------------------------
    [Function("EnrollmentJoin")]
    public async Task<IActionResult> JoinAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "enrollment/join")] HttpRequest req,
        CancellationToken ct)
    {
        JoinRequest? body;
        try
        {
            body = await JsonSerializer.DeserializeAsync<JoinRequest>(req.Body,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }, ct);
        }
        catch
        {
            return new BadRequestObjectResult("Invalid JSON body.");
        }

        if (body is null || string.IsNullOrEmpty(body.AadUserId) || string.IsNullOrEmpty(body.UserPrincipalName))
            return new BadRequestObjectResult("AadUserId and UserPrincipalName are required.");

        var existing = await _sp.GetUserByAadIdAsync(body.AadUserId, ct);
        bool isNew = existing is null;

        var user = existing ?? new CepUser
        {
            AadUserId = body.AadUserId,
            UserPrincipalName = body.UserPrincipalName,
            DisplayName = body.DisplayName ?? body.UserPrincipalName,
            Email = body.Email ?? "",
            EnrollmentDate = DateTime.UtcNow,
        };

        // Update / re-activate fields
        user.IsActive = true;
        user.Department = body.Department ?? user.Department;
        user.Team = body.Team ?? user.Team;
        user.IsEngagementNudgesEnabled = body.IsEngagementNudgesEnabled ?? user.IsEngagementNudgesEnabled;
        if (isNew) user.EnrollmentDate = DateTime.UtcNow;

        await _sp.UpsertUserAsync(user, ct);
        _log.LogInformation("{Action} enrollment for {UserId}", isNew ? "New" : "Reactivated", body.AadUserId);

        if (isNew)
            await _notifier.SendWelcomeAsync(user, ct);

        return new OkObjectResult(new { message = isNew ? "Enrolled successfully." : "Re-enrolled successfully.", userId = user.AadUserId });
    }

    // ------------------------------------------------------------------
    // POST /api/enrollment/leave
    // ------------------------------------------------------------------
    [Function("EnrollmentLeave")]
    public async Task<IActionResult> LeaveAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "enrollment/leave")] HttpRequest req,
        CancellationToken ct)
    {
        LeaveRequest? body;
        try
        {
            body = await JsonSerializer.DeserializeAsync<LeaveRequest>(req.Body,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }, ct);
        }
        catch
        {
            return new BadRequestObjectResult("Invalid JSON body.");
        }

        if (body is null || string.IsNullOrEmpty(body.AadUserId))
            return new BadRequestObjectResult("AadUserId is required.");

        var user = await _sp.GetUserByAadIdAsync(body.AadUserId, ct);
        if (user is null)
            return new NotFoundObjectResult("User not enrolled.");

        user.IsActive = false;
        await _sp.UpsertUserAsync(user, ct);
        _log.LogInformation("User {UserId} left the program", body.AadUserId);

        return new OkObjectResult(new { message = "Left the program successfully." });
    }

    // ------------------------------------------------------------------
    // DTOs
    // ------------------------------------------------------------------

    private record JoinRequest(
        string AadUserId,
        string UserPrincipalName,
        string? DisplayName,
        string? Email,
        string? Department,
        string? Team,
        bool? IsEngagementNudgesEnabled);

    private record LeaveRequest(string AadUserId);
}
