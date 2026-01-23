using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Formula;

/// <summary>
///     Result for getting formulas from Excel workbooks.
/// </summary>
public record GetFormulasResult
{
    /// <summary>
    ///     Number of formulas found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of formula information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<FormulaInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no formulas found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
