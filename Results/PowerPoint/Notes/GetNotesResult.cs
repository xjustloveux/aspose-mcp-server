using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Notes;

/// <summary>
///     Result for getting notes from PowerPoint presentations.
/// </summary>
public record GetNotesResult
{
    /// <summary>
    ///     Number of slides (when getting all notes).
    /// </summary>
    [JsonPropertyName("count")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Count { get; init; }

    /// <summary>
    ///     Slide index (when getting single slide notes).
    /// </summary>
    [JsonPropertyName("slideIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SlideIndex { get; init; }

    /// <summary>
    ///     Whether slide has notes (when getting single slide).
    /// </summary>
    [JsonPropertyName("hasNotes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? HasNotes { get; init; }

    /// <summary>
    ///     Notes text (when getting single slide).
    /// </summary>
    [JsonPropertyName("notes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Notes { get; init; }

    /// <summary>
    ///     List of slide notes (when getting all).
    /// </summary>
    [JsonPropertyName("slides")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<SlideNotesInfo>? Slides { get; init; }
}
