using System.Text.Json.Serialization;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Results.PowerPoint.Shape;

/// <summary>
///     Result type for getting detailed shape information from PowerPoint presentations.
/// </summary>
public sealed record GetShapeDetailsResult : GetShapeInfo
{
    /// <summary>
    ///     Gets whether the shape is hidden.
    /// </summary>
    [JsonPropertyName("hidden")]
    public required bool Hidden { get; init; }

    /// <summary>
    ///     Gets the alternative text for the shape.
    /// </summary>
    [JsonPropertyName("alternativeText")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AlternativeText { get; init; }

    /// <summary>
    ///     Gets whether the shape is flipped horizontally.
    /// </summary>
    [JsonPropertyName("flipHorizontal")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? FlipHorizontal { get; init; }

    /// <summary>
    ///     Gets whether the shape is flipped vertically.
    /// </summary>
    [JsonPropertyName("flipVertical")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? FlipVertical { get; init; }

    /// <summary>
    ///     Gets the shape-specific detail information.
    /// </summary>
    [JsonPropertyName("details")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ShapeDetails? Details { get; init; }
}
