using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Ocr;

/// <summary>
///     Bounding box coordinates for recognized text.
/// </summary>
public record OcrBoundingBox
{
    /// <summary>
    ///     X coordinate of the top-left corner.
    /// </summary>
    [JsonPropertyName("x")]
    public int X { get; init; }

    /// <summary>
    ///     Y coordinate of the top-left corner.
    /// </summary>
    [JsonPropertyName("y")]
    public int Y { get; init; }

    /// <summary>
    ///     Width of the bounding box.
    /// </summary>
    [JsonPropertyName("width")]
    public int Width { get; init; }

    /// <summary>
    ///     Height of the bounding box.
    /// </summary>
    [JsonPropertyName("height")]
    public int Height { get; init; }
}
