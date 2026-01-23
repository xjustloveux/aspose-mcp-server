using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Text;

/// <summary>
///     Result for text search operations in Word documents.
/// </summary>
public record TextSearchResult
{
    /// <summary>
    ///     The text or pattern that was searched for.
    /// </summary>
    [JsonPropertyName("searchText")]
    public required string SearchText { get; init; }

    /// <summary>
    ///     Whether regex was used for the search.
    /// </summary>
    [JsonPropertyName("useRegex")]
    public required bool UseRegex { get; init; }

    /// <summary>
    ///     Whether the search was case sensitive.
    /// </summary>
    [JsonPropertyName("caseSensitive")]
    public required bool CaseSensitive { get; init; }

    /// <summary>
    ///     Total number of matches found.
    /// </summary>
    [JsonPropertyName("matchCount")]
    public required int MatchCount { get; init; }

    /// <summary>
    ///     Whether the results were limited by maxResults.
    /// </summary>
    [JsonPropertyName("limitReached")]
    public required bool LimitReached { get; init; }

    /// <summary>
    ///     List of matches found.
    /// </summary>
    [JsonPropertyName("matches")]
    public required IReadOnlyList<TextSearchMatch> Matches { get; init; }
}
