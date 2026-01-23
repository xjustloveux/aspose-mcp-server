using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results;

/// <summary>
///     Wrapper for finalized tool results containing both data and output info.
/// </summary>
/// <typeparam name="T">The handler result type.</typeparam>
public sealed record FinalizedResult<T>
{
    /// <summary>
    ///     The handler result data.
    /// </summary>
    [JsonPropertyName("data")]
    public required T Data { get; init; }

    /// <summary>
    ///     Output information (path or session).
    /// </summary>
    [JsonPropertyName("output")]
    public required OutputInfo Output { get; init; }
}
