using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Shape;

/// <summary>
///     Result type for getting detailed shape information from PowerPoint presentations.
/// </summary>
public sealed record GetShapeDetailsResult
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

    /// <summary>
    ///     Gets whether the shape is hidden.
    /// </summary>
    [JsonPropertyName("hidden")]
    public required bool Hidden { get; init; }

    /// <summary>
    ///     Gets the auto shape type (for AutoShape).
    /// </summary>
    [JsonPropertyName("shapeType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShapeType { get; init; }

    /// <summary>
    ///     Gets the text content (for AutoShape).
    /// </summary>
    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; init; }

    /// <summary>
    ///     Gets the fill type (for AutoShape).
    /// </summary>
    [JsonPropertyName("fillType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FillType { get; init; }

    /// <summary>
    ///     Gets the row count (for Table).
    /// </summary>
    [JsonPropertyName("rows")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Rows { get; init; }

    /// <summary>
    ///     Gets the column count (for Table).
    /// </summary>
    [JsonPropertyName("columns")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Columns { get; init; }

    /// <summary>
    ///     Gets the shape count (for GroupShape).
    /// </summary>
    [JsonPropertyName("shapeCount")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ShapeCount { get; init; }
}
