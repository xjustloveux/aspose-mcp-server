using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Layout;

/// <summary>
///     Information about a master slide.
/// </summary>
public sealed record GetMasterInfo
{
    /// <summary>
    ///     Gets the master index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the master name.
    /// </summary>
    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; init; }

    /// <summary>
    ///     Gets the layout count.
    /// </summary>
    [JsonPropertyName("layoutCount")]
    public required int LayoutCount { get; init; }

    /// <summary>
    ///     Gets the layouts.
    /// </summary>
    [JsonPropertyName("layouts")]
    public required List<GetLayoutInfo> Layouts { get; init; }
}
