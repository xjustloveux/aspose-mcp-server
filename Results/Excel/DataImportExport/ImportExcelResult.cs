using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataImportExport;

/// <summary>
///     Result for data import operations.
/// </summary>
public record ImportExcelResult
{
    /// <summary>
    ///     Number of rows imported.
    /// </summary>
    [JsonPropertyName("rowCount")]
    public required int RowCount { get; init; }

    /// <summary>
    ///     Number of columns imported.
    /// </summary>
    [JsonPropertyName("columnCount")]
    public required int ColumnCount { get; init; }

    /// <summary>
    ///     Starting cell of the import.
    /// </summary>
    [JsonPropertyName("startCell")]
    public required string StartCell { get; init; }

    /// <summary>
    ///     Result message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
