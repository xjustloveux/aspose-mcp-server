using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Background;

/// <summary>
///     Result for getting background from PowerPoint presentations.
/// </summary>
public record GetBackgroundResult
{
    /// <summary>
    ///     Slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Whether slide has background.
    /// </summary>
    [JsonPropertyName("hasBackground")]
    public required bool HasBackground { get; init; }

    /// <summary>
    ///     Fill type.
    /// </summary>
    [JsonPropertyName("fillType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FillType { get; init; }

    /// <summary>
    ///     Background color in hex.
    /// </summary>
    [JsonPropertyName("color")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; init; }

    /// <summary>
    ///     Background opacity.
    /// </summary>
    [JsonPropertyName("opacity")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Opacity { get; init; }

    /// <summary>
    ///     Whether background is picture fill.
    /// </summary>
    [JsonPropertyName("isPictureFill")]
    public required bool IsPictureFill { get; init; }
}
