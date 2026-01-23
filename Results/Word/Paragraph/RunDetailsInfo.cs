using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Run details information.
/// </summary>
public sealed record RunDetailsInfo
{
    /// <summary>
    ///     Total number of runs.
    /// </summary>
    [JsonPropertyName("total")]
    public required int Total { get; init; }

    /// <summary>
    ///     Number of runs displayed (max 10).
    /// </summary>
    [JsonPropertyName("displayed")]
    public required int Displayed { get; init; }

    /// <summary>
    ///     List of run details.
    /// </summary>
    [JsonPropertyName("details")]
    public required IReadOnlyList<RunDetailInfo> Details { get; init; }
}
