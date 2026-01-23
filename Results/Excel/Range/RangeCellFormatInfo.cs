using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Range;

/// <summary>
///     Basic format information for a cell.
/// </summary>
public record RangeCellFormatInfo
{
    /// <summary>
    ///     Font name.
    /// </summary>
    [JsonPropertyName("fontName")]
    public required string FontName { get; init; }

    /// <summary>
    ///     Font size.
    /// </summary>
    [JsonPropertyName("fontSize")]
    public required double FontSize { get; init; }
}
