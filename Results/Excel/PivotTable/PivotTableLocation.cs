using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.PivotTable;

/// <summary>
///     Pivot table location information.
/// </summary>
public record PivotTableLocation
{
    /// <summary>
    ///     Range string.
    /// </summary>
    [JsonPropertyName("range")]
    public required string Range { get; init; }

    /// <summary>
    ///     Start row index.
    /// </summary>
    [JsonPropertyName("startRow")]
    public required int StartRow { get; init; }

    /// <summary>
    ///     End row index.
    /// </summary>
    [JsonPropertyName("endRow")]
    public required int EndRow { get; init; }

    /// <summary>
    ///     Start column index.
    /// </summary>
    [JsonPropertyName("startColumn")]
    public required int StartColumn { get; init; }

    /// <summary>
    ///     End column index.
    /// </summary>
    [JsonPropertyName("endColumn")]
    public required int EndColumn { get; init; }
}
