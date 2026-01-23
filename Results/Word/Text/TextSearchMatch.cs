using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Text;

/// <summary>
///     Information about a single text search match.
/// </summary>
public record TextSearchMatch
{
    /// <summary>
    ///     The matched text.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Zero-based paragraph index where the match was found.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     Context text around the match with the match highlighted in brackets.
    /// </summary>
    [JsonPropertyName("context")]
    public required string Context { get; init; }
}
