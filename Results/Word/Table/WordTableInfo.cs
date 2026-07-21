using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Table;

/// <summary>
///     Information about a single table.
/// </summary>
public record WordTableInfo
{
    /// <summary>
    ///     Zero-based index of the table.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Zero-based index of the section containing the table. Lets callers disambiguate tables
    ///     when the list spans the whole document (no sectionIndex filter).
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    public required int SectionIndex { get; init; }

    /// <summary>
    ///     Number of rows.
    /// </summary>
    [JsonPropertyName("rows")]
    public required int Rows { get; init; }

    /// <summary>
    ///     Number of columns.
    /// </summary>
    [JsonPropertyName("columns")]
    public required int Columns { get; init; }

    /// <summary>
    ///     Text preceding the table.
    /// </summary>
    [JsonPropertyName("precedingText")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? PrecedingText { get; init; }
}
