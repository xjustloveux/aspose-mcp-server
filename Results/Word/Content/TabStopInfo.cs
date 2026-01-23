using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Content;

/// <summary>
///     Information about a single tab stop.
/// </summary>
public sealed record TabStopInfo
{
    /// <summary>
    ///     Tab stop position in points.
    /// </summary>
    [JsonPropertyName("position")]
    public required double Position { get; init; }

    /// <summary>
    ///     Tab stop alignment (Left, Center, Right, Decimal, Bar).
    /// </summary>
    [JsonPropertyName("alignment")]
    public required string Alignment { get; init; }
}
