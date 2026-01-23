using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.ConditionalFormatting;

/// <summary>
///     A single condition in conditional formatting.
/// </summary>
public record ConditionalFormattingCondition
{
    /// <summary>
    ///     Zero-based index of the condition.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Operator type.
    /// </summary>
    [JsonPropertyName("operatorType")]
    public required string OperatorType { get; init; }

    /// <summary>
    ///     First formula.
    /// </summary>
    [JsonPropertyName("formula1")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula1 { get; init; }

    /// <summary>
    ///     Second formula.
    /// </summary>
    [JsonPropertyName("formula2")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula2 { get; init; }

    /// <summary>
    ///     Foreground color.
    /// </summary>
    [JsonPropertyName("foregroundColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ForegroundColor { get; init; }

    /// <summary>
    ///     Background color.
    /// </summary>
    [JsonPropertyName("backgroundColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BackgroundColor { get; init; }
}
