using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Styles;

/// <summary>
///     Result for getting styles from Word documents.
/// </summary>
public sealed record GetWordStylesResult
{
    /// <summary>
    ///     Total number of styles.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Indicates whether built-in styles are included.
    /// </summary>
    [JsonPropertyName("includeBuiltIn")]
    public required bool IncludeBuiltIn { get; init; }

    /// <summary>
    ///     Optional note about which styles are shown.
    /// </summary>
    [JsonPropertyName("note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; init; }

    /// <summary>
    ///     List of paragraph styles.
    /// </summary>
    [JsonPropertyName("paragraphStyles")]
    public required IReadOnlyList<StyleInfo> ParagraphStyles { get; init; }
}
