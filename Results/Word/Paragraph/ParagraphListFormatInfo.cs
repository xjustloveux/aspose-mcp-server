using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Paragraph list format information.
/// </summary>
public sealed record ParagraphListFormatInfo
{
    /// <summary>
    ///     Indicates this is a list item.
    /// </summary>
    [JsonPropertyName("isListItem")]
    public required bool IsListItem { get; init; }

    /// <summary>
    ///     List level number (0-8).
    /// </summary>
    [JsonPropertyName("listLevel")]
    public required int ListLevel { get; init; }

    /// <summary>
    ///     List identifier.
    /// </summary>
    [JsonPropertyName("listId")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ListId { get; init; }
}
