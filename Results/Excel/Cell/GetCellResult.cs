using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Cell;

/// <summary>
///     Result for getting cell value from Excel workbooks.
/// </summary>
public record GetCellResult
{
    /// <summary>
    ///     Cell reference (e.g., "A1").
    /// </summary>
    [JsonPropertyName("cell")]
    public required string Cell { get; init; }

    /// <summary>
    ///     Cell value as string.
    /// </summary>
    [JsonPropertyName("value")]
    public required string Value { get; init; }

    /// <summary>
    ///     Cell value type (e.g., "String", "Numeric").
    /// </summary>
    [JsonPropertyName("valueType")]
    public required string ValueType { get; init; }

    /// <summary>
    ///     Cell formula if present.
    /// </summary>
    [JsonPropertyName("formula")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula { get; init; }

    /// <summary>
    ///     Cell format information if requested.
    /// </summary>
    [JsonPropertyName("format")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public GetCellFormatInfo? Format { get; init; }
}
