using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Shape;

/// <summary>
///     Result type for getting shapes from PowerPoint presentations.
/// </summary>
public sealed record GetShapesResult
{
    /// <summary>
    ///     Gets the slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Gets the count of shapes.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Gets the shapes list.
    /// </summary>
    [JsonPropertyName("shapes")]
    public required List<GetShapeInfo> Shapes { get; init; }
}
