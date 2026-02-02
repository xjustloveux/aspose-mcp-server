using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for GroupShape elements.
/// </summary>
public sealed record GroupShapeDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the number of child shapes in the group.
    /// </summary>
    [JsonPropertyName("childCount")]
    public required int ChildCount { get; init; }

    /// <summary>
    ///     Gets the child shape information.
    /// </summary>
    [JsonPropertyName("children")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<GroupChildShapeInfo>? Children { get; init; }
}

/// <summary>
///     Information about a child shape within a group.
/// </summary>
public sealed record GroupChildShapeInfo
{
    /// <summary>
    ///     Gets the child shape index within the group.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the child shape name.
    /// </summary>
    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; init; }

    /// <summary>
    ///     Gets the child shape type name.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Gets the child shape position.
    /// </summary>
    [JsonPropertyName("position")]
    public required ShapePositionInfo Position { get; init; }

    /// <summary>
    ///     Gets the child shape size.
    /// </summary>
    [JsonPropertyName("size")]
    public required ShapeSizeInfo Size { get; init; }
}

/// <summary>
///     Position information for a shape.
/// </summary>
public sealed record ShapePositionInfo
{
    /// <summary>
    ///     Gets the X coordinate.
    /// </summary>
    [JsonPropertyName("x")]
    public required float X { get; init; }

    /// <summary>
    ///     Gets the Y coordinate.
    /// </summary>
    [JsonPropertyName("y")]
    public required float Y { get; init; }
}

/// <summary>
///     Size information for a shape.
/// </summary>
public sealed record ShapeSizeInfo
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
