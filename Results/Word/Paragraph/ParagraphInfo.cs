using System.Text.Json.Serialization;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Information about a single paragraph, addressed in the unified paragraph-address space so the
///     coordinates can be fed straight back into any word_* operation.
/// </summary>
public record ParagraphInfo
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
    ///     Opaque stable handle for this paragraph, emitted only in session mode. Pass it back as
    ///     `handle` to address this exact paragraph in later calls even after intervening inserts or
    ///     deletes shift the paragraph indices.
    /// </summary>
    [JsonPropertyName("handle")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Handle { get; init; }

    /// <summary>
    ///     Human-readable location type (Body, Comment, TextBox, Shape).
    /// </summary>
    [JsonPropertyName("location")]
    public required string Location { get; init; }

    /// <summary>
    ///     Style name.
    /// </summary>
    [JsonPropertyName("style")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Style { get; init; }

    /// <summary>
    ///     Paragraph text (truncated if too long).
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Full text length.
    /// </summary>
    [JsonPropertyName("textLength")]
    public required int TextLength { get; init; }

    /// <summary>
    ///     Comment information if in a comment.
    /// </summary>
    [JsonPropertyName("commentInfo")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CommentInfo { get; init; }
}
