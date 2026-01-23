using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Background information for a slide.
/// </summary>
public sealed record GetSlideDetailsBackgroundInfo
{
    /// <summary>
    ///     Gets the fill type.
    /// </summary>
    [JsonPropertyName("fillType")]
    public required string FillType { get; init; }
}
