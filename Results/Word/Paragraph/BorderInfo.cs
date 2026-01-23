using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Border information.
/// </summary>
public sealed record BorderInfo
{
    /// <summary>
    ///     Line style (Single, Double, Dotted, etc.).
    /// </summary>
    [JsonPropertyName("lineStyle")]
    public required string LineStyle { get; init; }

    /// <summary>
    ///     Line width in points.
    /// </summary>
    [JsonPropertyName("lineWidth")]
    public required double LineWidth { get; init; }

    /// <summary>
    ///     Border color name.
    /// </summary>
    [JsonPropertyName("color")]
    public required string Color { get; init; }
}
