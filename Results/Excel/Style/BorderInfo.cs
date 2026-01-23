using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Style;

/// <summary>
///     Information about a single border.
/// </summary>
public record BorderInfo
{
    /// <summary>
    ///     Line style of the border.
    /// </summary>
    [JsonPropertyName("lineStyle")]
    public required string LineStyle { get; init; }

    /// <summary>
    ///     Color of the border.
    /// </summary>
    [JsonPropertyName("color")]
    public required string Color { get; init; }
}
