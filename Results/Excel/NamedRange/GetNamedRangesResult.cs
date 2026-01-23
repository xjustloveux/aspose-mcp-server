using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.NamedRange;

/// <summary>
///     Result for getting named ranges from Excel workbooks.
/// </summary>
public record GetNamedRangesResult
{
    /// <summary>
    ///     Number of named ranges.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of named range information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<NamedRangeInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no named ranges found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
