using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Range;

/// <summary>
///     Information about a single cell in the range.
/// </summary>
public record RangeCellInfo
{
    /// <summary>
    ///     Cell reference (e.g., "A1").
    /// </summary>
    [JsonPropertyName("cell")]
    public required string Cell { get; init; }

    /// <summary>
    ///     Display value of the cell.
    /// </summary>
    [JsonPropertyName("value")]
    public required string Value { get; init; }

    /// <summary>
    ///     Formula if the cell contains one and includeFormulas is true.
    /// </summary>
    [JsonPropertyName("formula")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula { get; init; }

    /// <summary>
    ///     Format information if includeFormat is true.
    /// </summary>
    [JsonPropertyName("format")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public RangeCellFormatInfo? Format { get; init; }
}
