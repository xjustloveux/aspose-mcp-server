using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Font;

/// <summary>
///     Result of getting fonts from a presentation.
/// </summary>
public record GetFontsPptResult
{
    /// <summary>
    ///     Total number of fonts found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Number of embedded fonts.
    /// </summary>
    [JsonPropertyName("embeddedCount")]
    public required int EmbeddedCount { get; init; }

    /// <summary>
    ///     The list of font information.
    /// </summary>
    [JsonPropertyName("items")]
    public required List<PptFontInfo> Items { get; init; }

    /// <summary>
    ///     Human-readable message describing the result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
