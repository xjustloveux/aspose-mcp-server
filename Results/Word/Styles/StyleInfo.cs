using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Styles;

/// <summary>
///     Information about a single style.
/// </summary>
public sealed record StyleInfo
{
    /// <summary>
    ///     Style name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Indicates whether this is a built-in style.
    /// </summary>
    [JsonPropertyName("builtIn")]
    public required bool BuiltIn { get; init; }

    /// <summary>
    ///     Name of the base style (if any).
    /// </summary>
    [JsonPropertyName("basedOn")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BasedOn { get; init; }

    /// <summary>
    ///     Font name (when ASCII and Far East fonts are the same).
    /// </summary>
    [JsonPropertyName("font")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Font { get; init; }

    /// <summary>
    ///     ASCII font name (when different from Far East).
    /// </summary>
    [JsonPropertyName("fontAscii")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontAscii { get; init; }

    /// <summary>
    ///     Far East font name (when different from ASCII).
    /// </summary>
    [JsonPropertyName("fontFarEast")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontFarEast { get; init; }

    /// <summary>
    ///     Font size in points.
    /// </summary>
    [JsonPropertyName("fontSize")]
    public required double FontSize { get; init; }

    /// <summary>
    ///     Indicates bold formatting.
    /// </summary>
    [JsonPropertyName("bold")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Bold { get; init; }

    /// <summary>
    ///     Indicates italic formatting.
    /// </summary>
    [JsonPropertyName("italic")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Italic { get; init; }

    /// <summary>
    ///     Paragraph alignment (Left, Center, Right, Justify).
    /// </summary>
    [JsonPropertyName("alignment")]
    public required string Alignment { get; init; }

    /// <summary>
    ///     Space before paragraph in points.
    /// </summary>
    [JsonPropertyName("spaceBefore")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double SpaceBefore { get; init; }

    /// <summary>
    ///     Space after paragraph in points.
    /// </summary>
    [JsonPropertyName("spaceAfter")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double SpaceAfter { get; init; }
}
