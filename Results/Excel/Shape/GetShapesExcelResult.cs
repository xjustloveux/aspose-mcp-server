using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Shape;

/// <summary>
///     Result for getting shapes from an Excel worksheet.
/// </summary>
public record GetShapesExcelResult
{
    /// <summary>
    ///     Number of shapes found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Sheet index.
    /// </summary>
    [JsonPropertyName("sheetIndex")]
    public required int SheetIndex { get; init; }

    /// <summary>
    ///     List of shape information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelShapeInfo> Items { get; init; }

    /// <summary>
    ///     Optional message.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
