using System.Text.Json.Serialization;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Results.Word.List;

/// <summary>
///     Information about a list paragraph, addressed in the unified paragraph-address space so the
///     coordinates can be fed straight back into any word_* operation.
/// </summary>
public record ListParagraphInfo
{
    /// <summary>
    ///     Story-relative paragraph index (0-based within its storyType/section). This is the value to
    ///     pass back as paragraphIndex to edit, format, delete, and other word_* operations.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     The story the paragraph belongs to (Body, Header, Footer, TextBox, Comment, Footnote, Endnote).
    /// </summary>
    [JsonPropertyName("storyType")]
    public string StoryType { get; init; } = StoryTypes.Body;

    /// <summary>
    ///     Section index for Body/Header/Footer stories.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    public int SectionIndex { get; init; }

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
    ///     Global document-order index across all stories (for cross-reference / disambiguation).
    /// </summary>
    [JsonPropertyName("documentOrderIndex")]
    public int DocumentOrderIndex { get; init; }

    /// <summary>
    ///     Preview of the paragraph content (truncated to 50 characters).
    /// </summary>
    [JsonPropertyName("contentPreview")]
    public required string ContentPreview { get; init; }

    /// <summary>
    ///     Indicates whether the paragraph is a list item.
    /// </summary>
    [JsonPropertyName("isListItem")]
    public required bool IsListItem { get; init; }

    /// <summary>
    ///     List level number (0-8).
    /// </summary>
    [JsonPropertyName("listLevel")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ListLevel { get; init; }

    /// <summary>
    ///     List identifier.
    /// </summary>
    [JsonPropertyName("listId")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ListId { get; init; }

    /// <summary>
    ///     Index of the item within the list.
    /// </summary>
    [JsonPropertyName("listItemIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ListItemIndex { get; init; }

    /// <summary>
    ///     List level format information.
    /// </summary>
    [JsonPropertyName("listLevelFormat")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ListLevelFormatInfo? ListLevelFormat { get; init; }
}
