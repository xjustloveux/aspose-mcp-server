using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Tab stop information for paragraph format.
/// </summary>
public sealed record ParagraphTabStopInfo
{
    /// <summary>
    ///     Tab stop position in points.
    /// </summary>
    [JsonPropertyName("position")]
    public required double Position { get; init; }

    /// <summary>
    ///     Tab stop alignment.
    /// </summary>
    [JsonPropertyName("alignment")]
    public required string Alignment { get; init; }

    /// <summary>
    ///     Tab leader type.
    /// </summary>
    [JsonPropertyName("leader")]
    public required string Leader { get; init; }
}
