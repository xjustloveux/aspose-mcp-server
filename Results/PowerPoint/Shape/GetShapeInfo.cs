using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Shape;

/// <summary>
///     Information about a shape.
/// </summary>
public record GetShapeInfo
{
    /// <summary>
    ///     Gets the shape index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the shape name.
    /// </summary>
    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; init; }

    /// <summary>
    ///     Gets the shape type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Gets the X position.
    /// </summary>
    [JsonPropertyName("x")]
    public required float X { get; init; }

    /// <summary>
    ///     Gets the Y position.
    /// </summary>
    [JsonPropertyName("y")]
    public required float Y { get; init; }

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

    /// <summary>
    ///     Gets the rotation.
    /// </summary>
    [JsonPropertyName("rotation")]
    public required float Rotation { get; init; }
}
