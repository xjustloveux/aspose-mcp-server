using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Cell;

/// <summary>
///     Format information for a cell in GetCellResult.
/// </summary>
public record GetCellFormatInfo
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

    /// <summary>
    ///     Whether font is bold.
    /// </summary>
    [JsonPropertyName("bold")]
    public required bool Bold { get; init; }

    /// <summary>
    ///     Whether font is italic.
    /// </summary>
    [JsonPropertyName("italic")]
    public required bool Italic { get; init; }

    /// <summary>
    ///     Background color.
    /// </summary>
    [JsonPropertyName("backgroundColor")]
    public required string BackgroundColor { get; init; }

    /// <summary>
    ///     Number format.
    /// </summary>
    [JsonPropertyName("numberFormat")]
    public required int NumberFormat { get; init; }
}
