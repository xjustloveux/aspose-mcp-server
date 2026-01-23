using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Layout;

/// <summary>
///     Result type for getting layouts from PowerPoint presentations (single master).
/// </summary>
public sealed record GetLayoutsResult
{
    /// <summary>
    ///     Gets the master index (when filtering by master).
    /// </summary>
    [JsonPropertyName("masterIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? MasterIndex { get; init; }

    /// <summary>
    ///     Gets the layout count (when filtering by master).
    /// </summary>
    [JsonPropertyName("count")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Count { get; init; }

    /// <summary>
    ///     Gets the layouts list (when filtering by master).
    /// </summary>
    [JsonPropertyName("layouts")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<GetLayoutInfo>? Layouts { get; init; }

    /// <summary>
    ///     Gets the masters count (when getting all).
    /// </summary>
    [JsonPropertyName("mastersCount")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? MastersCount { get; init; }

    /// <summary>
    ///     Gets the masters list (when getting all).
    /// </summary>
    [JsonPropertyName("masters")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<GetLayoutMasterInfo>? Masters { get; init; }
}
