using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Hyperlink;

/// <summary>
///     Result for getting hyperlinks from PowerPoint presentations.
/// </summary>
public record GetHyperlinksPptResult
{
    /// <summary>
    ///     Slide index (when getting single slide).
    /// </summary>
    [JsonPropertyName("slideIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SlideIndex { get; init; }

    /// <summary>
    ///     Number of hyperlinks on slide (when getting single slide).
    /// </summary>
    [JsonPropertyName("count")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Count { get; init; }

    /// <summary>
    ///     Total count across all slides (when getting all slides).
    /// </summary>
    [JsonPropertyName("totalCount")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? TotalCount { get; init; }

    /// <summary>
    ///     List of hyperlinks (when getting single slide).
    /// </summary>
    [JsonPropertyName("hyperlinks")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<object>? Hyperlinks { get; init; }

    /// <summary>
    ///     List of slides with hyperlinks (when getting all slides).
    /// </summary>
    [JsonPropertyName("slides")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<SlideHyperlinksInfo>? Slides { get; init; }
}
