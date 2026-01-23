using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Filter settings for paragraph retrieval.
/// </summary>
public record ParagraphFilters
{
    /// <summary>
    ///     Section index filter.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SectionIndex { get; init; }

    /// <summary>
    ///     Whether empty paragraphs are included.
    /// </summary>
    [JsonPropertyName("includeEmpty")]
    public required bool IncludeEmpty { get; init; }

    /// <summary>
    ///     Style filter.
    /// </summary>
    [JsonPropertyName("styleFilter")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StyleFilter { get; init; }

    /// <summary>
    ///     Whether comment paragraphs are included.
    /// </summary>
    [JsonPropertyName("includeCommentParagraphs")]
    public required bool IncludeCommentParagraphs { get; init; }

    /// <summary>
    ///     Whether textbox paragraphs are included.
    /// </summary>
    [JsonPropertyName("includeTextboxParagraphs")]
    public required bool IncludeTextboxParagraphs { get; init; }
}
