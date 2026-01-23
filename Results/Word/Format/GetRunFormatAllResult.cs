using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Format;

/// <summary>
///     Result for getting run format information for all runs in a paragraph.
/// </summary>
public sealed record GetRunFormatAllResult
{
    /// <summary>
    ///     Zero-based paragraph index.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     Total number of runs.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of run format information.
    /// </summary>
    [JsonPropertyName("runs")]
    public required IReadOnlyList<RunFormatInfo> Runs { get; init; }
}
