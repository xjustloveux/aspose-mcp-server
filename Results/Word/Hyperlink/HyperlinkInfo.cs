using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Hyperlink;

/// <summary>
///     Information about a single hyperlink. When the hyperlink resolves to a paragraph, that
///     paragraph is reported in the unified paragraph-address space so the coordinates can be fed
///     straight back into any word_* operation.
/// </summary>
public record HyperlinkInfo
{
    /// <summary>
    ///     Zero-based index of the hyperlink.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Display text of the hyperlink.
    /// </summary>
    [JsonPropertyName("displayText")]
    public required string DisplayText { get; init; }

    /// <summary>
    ///     URL or address of the hyperlink.
    /// </summary>
    [JsonPropertyName("address")]
    public required string Address { get; init; }

    /// <summary>
    ///     Sub-address (e.g., bookmark within document).
    /// </summary>
    [JsonPropertyName("subAddress")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SubAddress { get; init; }

    /// <summary>
    ///     Tooltip text for the hyperlink.
    /// </summary>
    [JsonPropertyName("tooltip")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Tooltip { get; init; }

    /// <summary>
    ///     Story-relative index of the paragraph containing the hyperlink (0-based within its
    ///     storyType/section); null when the containing paragraph cannot be determined.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ParagraphIndex { get; init; }

    /// <summary>
    ///     The story the containing paragraph belongs to (Body, Header, Footer, TextBox, Comment,
    ///     Footnote, Endnote); null when the containing paragraph cannot be determined.
    /// </summary>
    [JsonPropertyName("storyType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoryType { get; init; }

    /// <summary>
    ///     Section index for Body/Header/Footer stories; null when the containing paragraph cannot be determined.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SectionIndex { get; init; }

    /// <summary>
    ///     For Header/Footer stories: Primary, First, or Even; null otherwise.
    /// </summary>
    [JsonPropertyName("headerFooterType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HeaderFooterType { get; init; }

    /// <summary>
    ///     Instance selector for multi-instance stories (e.g. Comment id); null when not applicable.
    /// </summary>
    [JsonPropertyName("containerIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ContainerIndex { get; init; }

    /// <summary>
    ///     Global document-order index of the containing paragraph; null when it cannot be determined.
    /// </summary>
    [JsonPropertyName("documentOrderIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? DocumentOrderIndex { get; init; }
}
