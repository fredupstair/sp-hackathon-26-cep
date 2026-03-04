namespace CepFunctions.Models;

/// <summary>
/// Message enqueued by the Orchestrator and processed by the Queue Worker.
/// </summary>
public class IngestMessage
{
    public string AadUserId { get; set; } = "";
    public string UserPrincipalName { get; set; } = "";
    public string UserEmail { get; set; } = "";

    /// <summary>Start of the Graph query window (exclusive lower bound).</summary>
    public DateTime WatermarkUtc { get; set; }

    /// <summary>End of the Graph query window (exclusive upper bound).</summary>
    public DateTime RunTimeUtc { get; set; }

    public string CorrelationId { get; set; } = "";
}
