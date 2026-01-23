using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Result type for getting slide details from PowerPoint presentations.
/// </summary>
public sealed record GetSlideDetailsResult
{
    /// <summary>
    ///     Gets the slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Gets whether the slide is hidden.
    /// </summary>
    [JsonPropertyName("hidden")]
    public required bool Hidden { get; init; }

    /// <summary>
    ///     Gets the layout name.
    /// </summary>
    [JsonPropertyName("layout")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Layout { get; init; }

    /// <summary>
    ///     Gets the slide size information.
    /// </summary>
    [JsonPropertyName("slideSize")]
    public required GetSlideDetailsSizeInfo SlideSize { get; init; }

    /// <summary>
    ///     Gets the shapes count.
    /// </summary>
    [JsonPropertyName("shapesCount")]
    public required int ShapesCount { get; init; }

    /// <summary>
    ///     Gets the transition information.
    /// </summary>
    [JsonPropertyName("transition")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public GetSlideDetailsTransitionInfo? Transition { get; init; }

    /// <summary>
    ///     Gets the animations count.
    /// </summary>
    [JsonPropertyName("animationsCount")]
    public required int AnimationsCount { get; init; }

    /// <summary>
    ///     Gets the animations list.
    /// </summary>
    [JsonPropertyName("animations")]
    public required List<GetSlideDetailsAnimationInfo> Animations { get; init; }

    /// <summary>
    ///     Gets the background information.
    /// </summary>
    [JsonPropertyName("background")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public GetSlideDetailsBackgroundInfo? Background { get; init; }

    /// <summary>
    ///     Gets the notes text.
    /// </summary>
    [JsonPropertyName("notes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Notes { get; init; }

    /// <summary>
    ///     Gets the thumbnail as base64.
    /// </summary>
    [JsonPropertyName("thumbnail")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Thumbnail { get; init; }
}
