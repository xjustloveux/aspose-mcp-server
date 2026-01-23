using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.List;

/// <summary>
///     Information about a list paragraph.
/// </summary>
public sealed record ListParagraphInfo
{
    /// <summary>
    ///     Zero-based paragraph index.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

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
