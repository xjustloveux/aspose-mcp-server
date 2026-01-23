using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataValidation;

/// <summary>
///     Result for getting data validations from Excel workbooks.
/// </summary>
public record GetDataValidationsResult
{
    /// <summary>
    ///     Number of data validations.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of data validation information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<DataValidationInfo> Items { get; init; }
}
