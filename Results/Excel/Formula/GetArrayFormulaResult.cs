using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Formula;

/// <summary>
///     Result for getting array formula information from an Excel cell.
/// </summary>
public record GetArrayFormulaResult
{
    /// <summary>
    ///     Cell reference.
    /// </summary>
    [JsonPropertyName("cell")]
    public required string Cell { get; init; }

    /// <summary>
    ///     Whether the cell contains an array formula.
    /// </summary>
    [JsonPropertyName("isArrayFormula")]
    public required bool IsArrayFormula { get; init; }

    /// <summary>
    ///     Formula in the cell (only present when isArrayFormula is true).
    /// </summary>
    [JsonPropertyName("formula")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula { get; init; }

    /// <summary>
    ///     Range of the array formula (only present when isArrayFormula is true).
    /// </summary>
    [JsonPropertyName("arrayRange")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ArrayRange { get; init; }

    /// <summary>
    ///     Message when no array formula found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
