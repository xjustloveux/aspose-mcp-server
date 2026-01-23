using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Style;

/// <summary>
///     Format information for a single cell.
/// </summary>
public record CellFormatInfo
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
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }

    /// <summary>
    ///     Cell formula if present.
    /// </summary>
    [JsonPropertyName("formula")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula { get; init; }

    /// <summary>
    ///     Cell data type.
    /// </summary>
    [JsonPropertyName("dataType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DataType { get; init; }

    /// <summary>
    ///     Format details.
    /// </summary>
    [JsonPropertyName("format")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public CellFormatDetails? Format { get; init; }
}
