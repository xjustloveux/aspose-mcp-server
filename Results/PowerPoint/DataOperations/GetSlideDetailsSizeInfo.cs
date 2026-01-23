using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Slide size information.
/// </summary>
public sealed record GetSlideDetailsSizeInfo
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
