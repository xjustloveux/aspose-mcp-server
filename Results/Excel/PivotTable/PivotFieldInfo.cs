using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.PivotTable;

/// <summary>
///     Pivot field information.
/// </summary>
public record PivotFieldInfo
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
}
