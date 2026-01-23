using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Temp file information DTO for result types.
/// </summary>
public record TempFileInfoDto
{
    /// <summary>
    ///     Session ID that created this temp file.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Original file path before disconnection.
    /// </summary>
    [JsonPropertyName("originalPath")]
    public required string OriginalPath { get; init; }

    /// <summary>
    ///     Document type.
    /// </summary>
    [JsonPropertyName("documentType")]
    public required string DocumentType { get; init; }

    /// <summary>
    ///     When the temp file was saved.
    /// </summary>
    [JsonPropertyName("savedAt")]
    public DateTime SavedAt { get; init; }

    /// <summary>
    ///     When the temp file expires.
    /// </summary>
    [JsonPropertyName("expiresAt")]
    public DateTime ExpiresAt { get; init; }

    /// <summary>
    ///     File size in MB.
    /// </summary>
    [JsonPropertyName("fileSizeMB")]
    public double FileSizeMb { get; init; }

    /// <summary>
    ///     Whether to prompt user on reconnect.
    /// </summary>
    [JsonPropertyName("promptOnReconnect")]
    public bool PromptOnReconnect { get; init; }
}
