using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataOperations;

/// <summary>
///     Result for find and replace operations in Excel worksheets.
/// </summary>
public record FindReplaceResult
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
    ///     Total number of replacements made.
    /// </summary>
    [JsonPropertyName("replacementCount")]
    public required int ReplacementCount { get; init; }
}
