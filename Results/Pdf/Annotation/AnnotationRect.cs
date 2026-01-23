using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Annotation;

/// <summary>
///     Rectangle for annotation position.
/// </summary>
public record AnnotationRect
{
    /// <summary>
    ///     X position.
    /// </summary>
    [JsonPropertyName("x")]
    public required double X { get; init; }

    /// <summary>
    ///     Y position.
    /// </summary>
    [JsonPropertyName("y")]
    public required double Y { get; init; }

    /// <summary>
    ///     Width.
    /// </summary>
    [JsonPropertyName("width")]
    public required double Width { get; init; }

    /// <summary>
    ///     Height.
    /// </summary>
    [JsonPropertyName("height")]
    public required double Height { get; init; }
}
