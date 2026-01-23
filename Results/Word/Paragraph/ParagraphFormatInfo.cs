using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Paragraph format information.
/// </summary>
public sealed record ParagraphFormatInfo
{
    /// <summary>
    ///     Style name.
    /// </summary>
    [JsonPropertyName("styleName")]
    public required string StyleName { get; init; }

    /// <summary>
    ///     Paragraph alignment (Left, Center, Right, Justify).
    /// </summary>
    [JsonPropertyName("alignment")]
    public required string Alignment { get; init; }

    /// <summary>
    ///     Left indent in points.
    /// </summary>
    [JsonPropertyName("leftIndent")]
    public required double LeftIndent { get; init; }

    /// <summary>
    ///     Right indent in points.
    /// </summary>
    [JsonPropertyName("rightIndent")]
    public required double RightIndent { get; init; }

    /// <summary>
    ///     First line indent in points.
    /// </summary>
    [JsonPropertyName("firstLineIndent")]
    public required double FirstLineIndent { get; init; }

    /// <summary>
    ///     Space before paragraph in points.
    /// </summary>
    [JsonPropertyName("spaceBefore")]
    public required double SpaceBefore { get; init; }

    /// <summary>
    ///     Space after paragraph in points.
    /// </summary>
    [JsonPropertyName("spaceAfter")]
    public required double SpaceAfter { get; init; }

    /// <summary>
    ///     Line spacing value.
    /// </summary>
    [JsonPropertyName("lineSpacing")]
    public required double LineSpacing { get; init; }

    /// <summary>
    ///     Line spacing rule (AtLeast, Exactly, Multiple, etc.).
    /// </summary>
    [JsonPropertyName("lineSpacingRule")]
    public required string LineSpacingRule { get; init; }
}
