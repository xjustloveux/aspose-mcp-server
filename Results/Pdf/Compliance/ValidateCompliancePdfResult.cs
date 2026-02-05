using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Compliance;

/// <summary>
///     Result for PDF compliance validation operation.
/// </summary>
public record ValidateCompliancePdfResult
{
    /// <summary>
    ///     Whether the document is compliant with the specified format.
    /// </summary>
    [JsonPropertyName("isCompliant")]
    public required bool IsCompliant { get; init; }

    /// <summary>
    ///     The compliance format checked (e.g., "PDF/A-1b").
    /// </summary>
    [JsonPropertyName("format")]
    public required string Format { get; init; }

    /// <summary>
    ///     Number of compliance errors found.
    /// </summary>
    [JsonPropertyName("errorCount")]
    public required int ErrorCount { get; init; }

    /// <summary>
    ///     Path to the validation log file.
    /// </summary>
    [JsonPropertyName("logPath")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LogPath { get; init; }

    /// <summary>
    ///     Human-readable message describing the validation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
