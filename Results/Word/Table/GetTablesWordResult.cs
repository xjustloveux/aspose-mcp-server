using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Table;

/// <summary>
///     Result for getting tables from Word documents.
/// </summary>
public record GetTablesWordResult
{
    /// <summary>
    ///     Number of tables found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Section index if specified.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SectionIndex { get; init; }

    /// <summary>
    ///     List of table information.
    /// </summary>
    [JsonPropertyName("tables")]
    public required IReadOnlyList<WordTableInfo> Tables { get; init; }
}
