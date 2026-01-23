using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Notes;

/// <summary>
///     Notes information for a single slide.
/// </summary>
public record SlideNotesInfo
{
    /// <summary>
    ///     Slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Whether slide has notes.
    /// </summary>
    [JsonPropertyName("hasNotes")]
    public required bool HasNotes { get; init; }

    /// <summary>
    ///     Notes text.
    /// </summary>
    [JsonPropertyName("notes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Notes { get; init; }
}
