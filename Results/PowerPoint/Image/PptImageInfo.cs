using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Image;

/// <summary>
///     Information about a single PowerPoint image.
/// </summary>
public record PptImageInfo
{
    /// <summary>
    ///     Zero-based index of the image.
    /// </summary>
    [JsonPropertyName("imageIndex")]
    public required int ImageIndex { get; init; }

    /// <summary>
    ///     X position.
    /// </summary>
    [JsonPropertyName("x")]
    public required float X { get; init; }

    /// <summary>
    ///     Y position.
    /// </summary>
    [JsonPropertyName("y")]
    public required float Y { get; init; }

    /// <summary>
    ///     Image width.
    /// </summary>
    [JsonPropertyName("width")]
    public required float Width { get; init; }

    /// <summary>
    ///     Image height.
    /// </summary>
    [JsonPropertyName("height")]
    public required float Height { get; init; }

    /// <summary>
    ///     Content type.
    /// </summary>
    [JsonPropertyName("contentType")]
    public required string ContentType { get; init; }
}
