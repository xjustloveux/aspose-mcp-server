using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Layout;

/// <summary>
///     Master information with its layouts.
/// </summary>
public sealed record GetLayoutMasterInfo
{
    /// <summary>
    ///     Gets the master index.
    /// </summary>
    [JsonPropertyName("masterIndex")]
    public required int MasterIndex { get; init; }

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
