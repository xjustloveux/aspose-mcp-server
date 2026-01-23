using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Formula;

/// <summary>
///     Information about a single formula.
/// </summary>
public record FormulaInfo
{
    /// <summary>
    ///     Cell reference.
    /// </summary>
    [JsonPropertyName("cell")]
    public required string Cell { get; init; }

    /// <summary>
    ///     Formula string.
    /// </summary>
    [JsonPropertyName("formula")]
    public required string Formula { get; init; }

    /// <summary>
    ///     Calculated value.
    /// </summary>
    [JsonPropertyName("value")]
    public required string Value { get; init; }
}
