using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.NamedRange;

/// <summary>
///     Information about a single named range.
/// </summary>
public record NamedRangeInfo
{
    /// <summary>
    ///     Zero-based index of the named range.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Name of the named range.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Reference formula.
    /// </summary>
    [JsonPropertyName("reference")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Reference { get; init; }

    /// <summary>
    ///     Comment.
    /// </summary>
    [JsonPropertyName("comment")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Comment { get; init; }

    /// <summary>
    ///     Whether the named range is visible.
    /// </summary>
    [JsonPropertyName("isVisible")]
    public required bool IsVisible { get; init; }
}
