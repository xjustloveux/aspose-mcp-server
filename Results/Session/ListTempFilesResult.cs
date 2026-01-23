using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for list temp files operation.
/// </summary>
public record ListTempFilesResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Number of recoverable files.
    /// </summary>
    [JsonPropertyName("count")]
    public int Count { get; init; }

    /// <summary>
    ///     List of recoverable files.
    /// </summary>
    [JsonPropertyName("files")]
    public required IReadOnlyList<TempFileInfoDto> Files { get; init; }
}
