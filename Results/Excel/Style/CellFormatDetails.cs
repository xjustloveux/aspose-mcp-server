using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Style;

/// <summary>
///     Detailed format information for a cell.
/// </summary>
public record CellFormatDetails
{
    /// <summary>
    ///     Font name.
    /// </summary>
    [JsonPropertyName("fontName")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontName { get; init; }

    /// <summary>
    ///     Font size.
    /// </summary>
    [JsonPropertyName("fontSize")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? FontSize { get; init; }

    /// <summary>
    ///     Whether font is bold.
    /// </summary>
    [JsonPropertyName("bold")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Bold { get; init; }

    /// <summary>
    ///     Whether font is italic.
    /// </summary>
    [JsonPropertyName("italic")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Italic { get; init; }

    /// <summary>
    ///     Underline style.
    /// </summary>
    [JsonPropertyName("underline")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Underline { get; init; }

    /// <summary>
    ///     Whether font has strikethrough.
    /// </summary>
    [JsonPropertyName("strikethrough")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Strikethrough { get; init; }

    /// <summary>
    ///     Font color.
    /// </summary>
    [JsonPropertyName("fontColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontColor { get; init; }

    /// <summary>
    ///     Foreground color.
    /// </summary>
    [JsonPropertyName("foregroundColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ForegroundColor { get; init; }

    /// <summary>
    ///     Background color.
    /// </summary>
    [JsonPropertyName("backgroundColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BackgroundColor { get; init; }

    /// <summary>
    ///     Pattern type.
    /// </summary>
    [JsonPropertyName("patternType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? PatternType { get; init; }

    /// <summary>
    ///     Horizontal alignment.
    /// </summary>
    [JsonPropertyName("horizontalAlignment")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HorizontalAlignment { get; init; }

    /// <summary>
    ///     Vertical alignment.
    /// </summary>
    [JsonPropertyName("verticalAlignment")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? VerticalAlignment { get; init; }

    /// <summary>
    ///     Number format.
    /// </summary>
    [JsonPropertyName("numberFormat")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? NumberFormat { get; init; }

    /// <summary>
    ///     Custom format string.
    /// </summary>
    [JsonPropertyName("customFormat")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CustomFormat { get; init; }

    /// <summary>
    ///     Border information.
    /// </summary>
    [JsonPropertyName("borders")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public BordersInfo? Borders { get; init; }
}
