using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Slide;

/// <summary>
///     Information about a slide.
/// </summary>
public sealed record GetSlideInfoItem
{
    /// <summary>
    ///     Gets the slide index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the slide title.
    /// </summary>
    [JsonPropertyName("title")]
    public required string Title { get; init; }

    /// <summary>
    ///     Gets the layout type.
    /// </summary>
    [JsonPropertyName("layoutType")]
    public required string LayoutType { get; init; }

    /// <summary>
    ///     Gets the layout name.
    /// </summary>
    [JsonPropertyName("layoutName")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LayoutName { get; init; }

    /// <summary>
    ///     Gets the shapes count.
    /// </summary>
    [JsonPropertyName("shapesCount")]
    public required int ShapesCount { get; init; }

    /// <summary>
    ///     Gets whether the slide has speaker notes.
    /// </summary>
    [JsonPropertyName("hasSpeakerNotes")]
    public required bool HasSpeakerNotes { get; init; }

    /// <summary>
    ///     Gets whether the slide is hidden.
    /// </summary>
    [JsonPropertyName("hidden")]
    public required bool Hidden { get; init; }
}
