using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Watermark;

/// <summary>
///     Result of getting watermarks from a presentation.
/// </summary>
public record GetWatermarksPptResult
{
    /// <summary>
    ///     Total number of watermark shapes found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     The list of watermark information.
    /// </summary>
    [JsonPropertyName("items")]
    public required List<PptWatermarkInfo> Items { get; init; }

    /// <summary>
    ///     Human-readable message describing the result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
