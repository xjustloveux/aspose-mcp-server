using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Text;

/// <summary>
///     Result for text replace operations in PowerPoint presentations.
/// </summary>
public record TextReplaceResult
{
    /// <summary>
    ///     The text that was searched for.
    /// </summary>
    [JsonPropertyName("findText")]
    public required string FindText { get; init; }

    /// <summary>
    ///     The text used for replacement.
    /// </summary>
    [JsonPropertyName("replaceText")]
    public required string ReplaceText { get; init; }

    /// <summary>
    ///     Total number of occurrences replaced.
    /// </summary>
    [JsonPropertyName("replacementCount")]
    public required int ReplacementCount { get; init; }
}
