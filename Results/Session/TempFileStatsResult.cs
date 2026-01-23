using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for temp file stats operation.
/// </summary>
public record TempFileStatsResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Total number of temp files.
    /// </summary>
    [JsonPropertyName("totalCount")]
    public int TotalCount { get; init; }

    /// <summary>
    ///     Total size of all temp files in MB.
    /// </summary>
    [JsonPropertyName("totalSizeMB")]
    public double TotalSizeMb { get; init; }

    /// <summary>
    ///     Number of expired temp files.
    /// </summary>
    [JsonPropertyName("expiredCount")]
    public int ExpiredCount { get; init; }

    /// <summary>
    ///     Retention hours configuration.
    /// </summary>
    [JsonPropertyName("retentionHours")]
    public int RetentionHours { get; init; }
}
