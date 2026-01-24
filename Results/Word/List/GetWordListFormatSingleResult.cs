using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.List;

/// <summary>
///     Result for getting list format information for a single paragraph.
/// </summary>
public sealed record GetWordListFormatSingleResult : ListParagraphInfo
{
    /// <summary>
    ///     Note for non-list paragraphs.
    /// </summary>
    [JsonPropertyName("note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; init; }
}
