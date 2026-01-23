using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Link;

/// <summary>
///     Information about a single link annotation.
/// </summary>
public record LinkInfo
{
    /// <summary>
    ///     Zero-based index of the link within the page.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     One-based page index where the link is located.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    public required int PageIndex { get; init; }

    /// <summary>
    ///     X coordinate (lower-left) of the link rectangle.
    /// </summary>
    [JsonPropertyName("x")]
    public required double X { get; init; }

    /// <summary>
    ///     Y coordinate (lower-left) of the link rectangle.
    /// </summary>
    [JsonPropertyName("y")]
    public required double Y { get; init; }

    /// <summary>
    ///     Type of link ("url" or "page").
    /// </summary>
    [JsonPropertyName("type")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Type { get; init; }

    /// <summary>
    ///     URL for external links.
    /// </summary>
    [JsonPropertyName("url")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Url { get; init; }

    /// <summary>
    ///     Destination page number for internal links.
    /// </summary>
    [JsonPropertyName("destinationPage")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? DestinationPage { get; init; }
}
