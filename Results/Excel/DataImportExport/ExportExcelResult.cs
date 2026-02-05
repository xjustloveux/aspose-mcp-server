using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataImportExport;

/// <summary>
///     Result for data export operations.
/// </summary>
public record ExportExcelResult
{
    /// <summary>
    ///     Output file path.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }

    /// <summary>
    ///     Number of rows exported.
    /// </summary>
    [JsonPropertyName("rowCount")]
    public int RowCount { get; init; }

    /// <summary>
    ///     Result message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
