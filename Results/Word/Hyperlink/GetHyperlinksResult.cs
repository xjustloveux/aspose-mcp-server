using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Hyperlink;

/// <summary>
///     Result for getting hyperlinks from Word documents.
/// </summary>
public record GetHyperlinksResult
{
    /// <summary>
    ///     Total number of hyperlinks in the document.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of hyperlink information.
    /// </summary>
    [JsonPropertyName("hyperlinks")]
    public required IReadOnlyList<HyperlinkInfo> Hyperlinks { get; init; }

    /// <summary>
    ///     Optional message when no hyperlinks found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
