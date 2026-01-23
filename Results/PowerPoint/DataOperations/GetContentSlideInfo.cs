using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Information about a slide's content.
/// </summary>
public sealed record GetContentSlideInfo
{
    /// <summary>
    ///     Gets the slide index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets whether the slide is hidden.
    /// </summary>
    [JsonPropertyName("hidden")]
    public required bool Hidden { get; init; }

    /// <summary>
    ///     Gets the text content from the slide.
    /// </summary>
    [JsonPropertyName("textContent")]
    public required List<string> TextContent { get; init; }
}
