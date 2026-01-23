using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Information about the used range.
/// </summary>
public record UsedRangeInfo
{
    /// <summary>
    ///     Number of rows in the used range.
    /// </summary>
    [JsonPropertyName("rowCount")]
    public required int RowCount { get; init; }

    /// <summary>
    ///     Number of columns in the used range.
    /// </summary>
    [JsonPropertyName("columnCount")]
    public required int ColumnCount { get; init; }
}
