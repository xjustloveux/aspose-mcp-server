using System.Text.Json.Serialization;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Results.Word.Text;

/// <summary>
///     Information about a single text search match, addressed in the unified paragraph-address
///     space so the coordinates can be fed straight back into any word_* operation.
/// </summary>
public record TextSearchMatch
{
    /// <summary>
    ///     The matched text.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Story-relative paragraph index of the match (0-based within its storyType/section).
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     The story the matched paragraph belongs to (Body, Header, Footer, TextBox, Comment, Footnote, Endnote).
    /// </summary>
    [JsonPropertyName("storyType")]
    public string StoryType { get; init; } = StoryTypes.Body;

    /// <summary>
    ///     Section index for Body/Header/Footer/TextBox stories.
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
    ///     Opaque stable handle for the matched paragraph, emitted only in session mode. Pass it back
    ///     as `handle` to address this exact paragraph in later calls even after intervening inserts
    ///     or deletes shift the paragraph indices.
    /// </summary>
    [JsonPropertyName("handle")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Handle { get; init; }

    /// <summary>
    ///     Context text around the match with the match highlighted in brackets.
    /// </summary>
    [JsonPropertyName("context")]
    public required string Context { get; init; }
}
