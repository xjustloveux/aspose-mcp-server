using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.List;

/// <summary>
///     List level format information.
/// </summary>
public sealed record ListLevelFormatInfo
{
    /// <summary>
    ///     The number format symbol.
    /// </summary>
    [JsonPropertyName("symbol")]
    public required string Symbol { get; init; }

    /// <summary>
    ///     Text alignment (Left, Center, Right).
    /// </summary>
    [JsonPropertyName("alignment")]
    public required string Alignment { get; init; }

    /// <summary>
    ///     Text position in points.
    /// </summary>
    [JsonPropertyName("textPosition")]
    public required double TextPosition { get; init; }

    /// <summary>
    ///     Number style (Arabic, LowercaseLetter, etc.).
    /// </summary>
    [JsonPropertyName("numberStyle")]
    public required string NumberStyle { get; init; }
}
