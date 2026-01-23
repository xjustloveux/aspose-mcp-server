using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for list sessions operation.
/// </summary>
public record ListSessionsResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Number of active sessions.
    /// </summary>
    [JsonPropertyName("count")]
    public int Count { get; init; }

    /// <summary>
    ///     Total memory usage in MB.
    /// </summary>
    [JsonPropertyName("totalMemoryMB")]
    public double TotalMemoryMb { get; init; }

    /// <summary>
    ///     List of active sessions.
    /// </summary>
    [JsonPropertyName("sessions")]
    public required IReadOnlyList<SessionInfoDto> Sessions { get; init; }
}
