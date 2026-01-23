using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.ConditionalFormatting;

/// <summary>
///     Information about a single conditional formatting.
/// </summary>
public record ConditionalFormattingInfo
{
    /// <summary>
    ///     Zero-based index of the conditional formatting.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Cell areas affected.
    /// </summary>
    [JsonPropertyName("areas")]
    public required IReadOnlyList<string> Areas { get; init; }

    /// <summary>
    ///     Number of conditions.
    /// </summary>
    [JsonPropertyName("conditionsCount")]
    public required int ConditionsCount { get; init; }

    /// <summary>
    ///     Conditions.
    /// </summary>
    [JsonPropertyName("conditions")]
    public required IReadOnlyList<ConditionalFormattingCondition> Conditions { get; init; }
}
