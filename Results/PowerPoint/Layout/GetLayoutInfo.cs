using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Layout;

/// <summary>
///     Layout information.
/// </summary>
public sealed record GetLayoutInfo
{
    /// <summary>
    ///     Gets the layout index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the layout name.
    /// </summary>
    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; init; }

    /// <summary>
    ///     Gets the layout type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }
}
