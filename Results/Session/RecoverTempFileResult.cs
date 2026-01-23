using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for recover temp file operation.
/// </summary>
public record RecoverTempFileResult
{
    /// <summary>
    ///     Whether recovery was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Session ID that was recovered.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Path where file was recovered to.
    /// </summary>
    [JsonPropertyName("recoveredPath")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RecoveredPath { get; init; }

    /// <summary>
    ///     Original file path before disconnection.
    /// </summary>
    [JsonPropertyName("originalPath")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? OriginalPath { get; init; }

    /// <summary>
    ///     Error message if recovery failed.
    /// </summary>
    [JsonPropertyName("errorMessage")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorMessage { get; init; }

    /// <summary>
    ///     Human-readable message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
