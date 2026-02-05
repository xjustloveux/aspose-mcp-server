using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Compliance;

/// <summary>
///     Result for PDF compliance conversion operation.
/// </summary>
public record ConvertCompliancePdfResult
{
    /// <summary>
    ///     The target compliance format converted to.
    /// </summary>
    [JsonPropertyName("format")]
    public required string Format { get; init; }

    /// <summary>
    ///     Whether the conversion was successful.
    /// </summary>
    [JsonPropertyName("isSuccess")]
    public required bool IsSuccess { get; init; }

    /// <summary>
    ///     Path to the conversion log file.
    /// </summary>
    [JsonPropertyName("logPath")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LogPath { get; init; }

    /// <summary>
    ///     Human-readable message describing the conversion result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
