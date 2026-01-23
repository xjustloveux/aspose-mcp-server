using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.PivotTable;

/// <summary>
///     Pivot data field information.
/// </summary>
public record PivotDataFieldInfo
{
    /// <summary>
    ///     Field name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Field position.
    /// </summary>
    [JsonPropertyName("position")]
    public required int Position { get; init; }

    /// <summary>
    ///     Aggregation function.
    /// </summary>
    [JsonPropertyName("function")]
    public required string Function { get; init; }
}
