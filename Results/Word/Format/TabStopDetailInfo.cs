using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Format;

/// <summary>
///     Detailed tab stop information.
/// </summary>
public sealed record TabStopDetailInfo
{
    /// <summary>
    ///     Tab stop position in points.
    /// </summary>
    [JsonPropertyName("positionPt")]
    public required double PositionPt { get; init; }

    /// <summary>
    ///     Tab stop position in centimeters.
    /// </summary>
    [JsonPropertyName("positionCm")]
    public required double PositionCm { get; init; }

    /// <summary>
    ///     Tab stop alignment (Left, Center, Right, Decimal, Bar).
    /// </summary>
    [JsonPropertyName("alignment")]
    public required string Alignment { get; init; }

    /// <summary>
    ///     Tab leader type (None, Dots, Dashes, Line, Heavy, MiddleDot).
    /// </summary>
    [JsonPropertyName("leader")]
    public required string Leader { get; init; }

    /// <summary>
    ///     Source of the tab stop (Paragraph, Style, etc.).
    /// </summary>
    [JsonPropertyName("source")]
    public required string Source { get; init; }
}
