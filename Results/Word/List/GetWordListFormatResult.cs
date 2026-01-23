using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.List;

/// <summary>
///     Result for getting list format information from Word documents.
/// </summary>
public sealed record GetWordListFormatResult
{
    /// <summary>
    ///     Total count of list paragraphs.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of list paragraph information.
    /// </summary>
    [JsonPropertyName("listParagraphs")]
    public required IReadOnlyList<ListParagraphInfo> ListParagraphs { get; init; }

    /// <summary>
    ///     Optional message when no list paragraphs are found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
