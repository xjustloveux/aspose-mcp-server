using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Section;

/// <summary>
///     Information about a single section.
/// </summary>
public record SectionInfo
{
    /// <summary>
    ///     Zero-based index of the section.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Section name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Start slide index.
    /// </summary>
    [JsonPropertyName("startSlideIndex")]
    public required int StartSlideIndex { get; init; }

    /// <summary>
    ///     Number of slides in section.
    /// </summary>
    [JsonPropertyName("slideCount")]
    public required int SlideCount { get; init; }
}
