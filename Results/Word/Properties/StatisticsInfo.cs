using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Properties;

/// <summary>
///     Document statistics.
/// </summary>
public sealed record StatisticsInfo
{
    /// <summary>
    ///     Word count.
    /// </summary>
    [JsonPropertyName("wordCount")]
    public required int WordCount { get; init; }

    /// <summary>
    ///     Character count.
    /// </summary>
    [JsonPropertyName("characterCount")]
    public required int CharacterCount { get; init; }

    /// <summary>
    ///     Page count.
    /// </summary>
    [JsonPropertyName("pageCount")]
    public required int PageCount { get; init; }

    /// <summary>
    ///     Paragraph count.
    /// </summary>
    [JsonPropertyName("paragraphCount")]
    public required int ParagraphCount { get; init; }

    /// <summary>
    ///     Line count.
    /// </summary>
    [JsonPropertyName("lineCount")]
    public required int LineCount { get; init; }
}
