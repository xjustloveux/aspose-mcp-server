using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataOperations;

/// <summary>
///     Statistics for a specific cell range.
/// </summary>
public record RangeStatistics
{
    /// <summary>
    ///     The range address.
    /// </summary>
    [JsonPropertyName("range")]
    public required string Range { get; init; }

    /// <summary>
    ///     Total number of cells in the range.
    /// </summary>
    [JsonPropertyName("totalCells")]
    public required int TotalCells { get; init; }

    /// <summary>
    ///     Number of cells containing numeric values.
    /// </summary>
    [JsonPropertyName("numericCells")]
    public required int NumericCells { get; init; }

    /// <summary>
    ///     Number of cells containing non-numeric values.
    /// </summary>
    [JsonPropertyName("nonNumericCells")]
    public required int NonNumericCells { get; init; }

    /// <summary>
    ///     Number of empty cells.
    /// </summary>
    [JsonPropertyName("emptyCells")]
    public required int EmptyCells { get; init; }

    /// <summary>
    ///     Sum of numeric values (only present when there are numeric cells).
    /// </summary>
    [JsonPropertyName("sum")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Sum { get; init; }

    /// <summary>
    ///     Average of numeric values (only present when there are numeric cells).
    /// </summary>
    [JsonPropertyName("average")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Average { get; init; }

    /// <summary>
    ///     Minimum numeric value (only present when there are numeric cells).
    /// </summary>
    [JsonPropertyName("min")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Min { get; init; }

    /// <summary>
    ///     Maximum numeric value (only present when there are numeric cells).
    /// </summary>
    [JsonPropertyName("max")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Max { get; init; }

    /// <summary>
    ///     Count of numeric values (only present when there are numeric cells).
    /// </summary>
    [JsonPropertyName("count")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Count { get; init; }
}
