using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for cleanup temp files operation.
/// </summary>
public record CleanupTempFilesResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Number of temp files scanned.
    /// </summary>
    [JsonPropertyName("scannedCount")]
    public int ScannedCount { get; init; }

    /// <summary>
    ///     Number of temp files deleted.
    /// </summary>
    [JsonPropertyName("deletedCount")]
    public int DeletedCount { get; init; }

    /// <summary>
    ///     Number of errors during cleanup.
    /// </summary>
    [JsonPropertyName("errorCount")]
    public int ErrorCount { get; init; }

    /// <summary>
    ///     Human-readable message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
