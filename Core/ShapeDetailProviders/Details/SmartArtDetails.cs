using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for SmartArt elements.
/// </summary>
public sealed record SmartArtDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the SmartArt layout type.
    /// </summary>
    [JsonPropertyName("layout")]
    public required string Layout { get; init; }

    /// <summary>
    ///     Gets the quick style applied to the SmartArt.
    /// </summary>
    [JsonPropertyName("quickStyle")]
    public required string QuickStyle { get; init; }

    /// <summary>
    ///     Gets the color style applied to the SmartArt.
    /// </summary>
    [JsonPropertyName("colorStyle")]
    public required string ColorStyle { get; init; }

    /// <summary>
    ///     Gets whether the SmartArt layout is reversed.
    /// </summary>
    [JsonPropertyName("isReversed")]
    public required bool IsReversed { get; init; }

    /// <summary>
    ///     Gets the total number of nodes.
    /// </summary>
    [JsonPropertyName("nodeCount")]
    public required int NodeCount { get; init; }

    /// <summary>
    ///     Gets the node hierarchy information.
    /// </summary>
    [JsonPropertyName("nodes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<SmartArtNodeInfo>? Nodes { get; init; }
}

/// <summary>
///     Information about a SmartArt node, including its children (recursive).
/// </summary>
public sealed record SmartArtNodeInfo
{
    /// <summary>
    ///     Gets the node text content.
    /// </summary>
    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; init; }

    /// <summary>
    ///     Gets the node level in the hierarchy.
    /// </summary>
    [JsonPropertyName("level")]
    public required int Level { get; init; }

    /// <summary>
    ///     Gets whether the node is hidden.
    /// </summary>
    [JsonPropertyName("isHidden")]
    public required bool IsHidden { get; init; }

    /// <summary>
    ///     Gets the number of child nodes.
    /// </summary>
    [JsonPropertyName("childCount")]
    public required int ChildCount { get; init; }

    /// <summary>
    ///     Gets the child node information (recursive).
    /// </summary>
    [JsonPropertyName("children")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<SmartArtNodeInfo>? Children { get; init; }
}
