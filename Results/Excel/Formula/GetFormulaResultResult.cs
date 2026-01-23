using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Formula;

/// <summary>
///     Result for getting a formula result from an Excel cell.
/// </summary>
public record GetFormulaResultResult
{
    /// <summary>
    ///     Cell reference.
    /// </summary>
    [JsonPropertyName("cell")]
    public required string Cell { get; init; }

    /// <summary>
    ///     Formula in the cell, or null if no formula.
    /// </summary>
    [JsonPropertyName("formula")]
    public string? Formula { get; init; }

    /// <summary>
    ///     Calculated value of the cell.
    /// </summary>
    [JsonPropertyName("calculatedValue")]
    public required string CalculatedValue { get; init; }

    /// <summary>
    ///     Type of the cell value.
    /// </summary>
    [JsonPropertyName("valueType")]
    public required string ValueType { get; init; }
}
