using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Image;

/// <summary>
///     Result for getting images from Excel workbooks.
/// </summary>
public record GetImagesExcelResult
{
    /// <summary>
    ///     Number of images found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of image information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelImageInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no images found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
