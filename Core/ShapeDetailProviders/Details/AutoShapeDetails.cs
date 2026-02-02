using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for AutoShape elements.
/// </summary>
public sealed record AutoShapeDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the auto shape type (e.g., Rectangle, Ellipse).
    /// </summary>
    [JsonPropertyName("shapeType")]
    public required string ShapeType { get; init; }

    /// <summary>
    ///     Gets the text content of the shape.
    /// </summary>
    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; init; }

    /// <summary>
    ///     Gets whether the shape has a text frame.
    /// </summary>
    [JsonPropertyName("hasTextFrame")]
    public required bool HasTextFrame { get; init; }

    /// <summary>
    ///     Gets the number of paragraphs in the text frame.
    /// </summary>
    [JsonPropertyName("paragraphCount")]
    public required int ParagraphCount { get; init; }

    /// <summary>
    ///     Gets the hyperlink target, if any.
    /// </summary>
    [JsonPropertyName("hyperlink")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Hyperlink { get; init; }

    /// <summary>
    ///     Gets the fill type (e.g., Solid, Gradient).
    /// </summary>
    [JsonPropertyName("fillType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FillType { get; init; }

    /// <summary>
    ///     Gets the solid fill color in #RRGGBB format. Only present when fill type is Solid.
    /// </summary>
    [JsonPropertyName("fillColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FillColor { get; init; }

    /// <summary>
    ///     Gets the fill transparency (0.0 = opaque, 1.0 = fully transparent). Only present when fill type is Solid.
    /// </summary>
    [JsonPropertyName("transparency")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public float? Transparency { get; init; }

    /// <summary>
    ///     Gets the line color in #RRGGBB format.
    /// </summary>
    [JsonPropertyName("lineColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LineColor { get; init; }

    /// <summary>
    ///     Gets the line width in points.
    /// </summary>
    [JsonPropertyName("lineWidth")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? LineWidth { get; init; }

    /// <summary>
    ///     Gets the line dash style (e.g., Solid, Dash, Dot).
    /// </summary>
    [JsonPropertyName("lineDashStyle")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LineDashStyle { get; init; }

    /// <summary>
    ///     Gets the shape adjustment values, if any.
    /// </summary>
    [JsonPropertyName("adjustments")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<AdjustmentInfo>? Adjustments { get; init; }
}

/// <summary>
///     Information about a shape adjustment value.
/// </summary>
public sealed record AdjustmentInfo
{
    /// <summary>
    ///     Gets the adjustment index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the raw adjustment value.
    /// </summary>
    [JsonPropertyName("value")]
    public required long Value { get; init; }
}
