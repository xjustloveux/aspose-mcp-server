using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Size information for statistics.
/// </summary>
public sealed record GetStatisticsSizeInfo
{
    /// <summary>
    ///     Gets the width.
    /// </summary>
    [JsonPropertyName("width")]
    public required float Width { get; init; }

    /// <summary>
    ///     Gets the height.
    /// </summary>
    [JsonPropertyName("height")]
    public required float Height { get; init; }
}
