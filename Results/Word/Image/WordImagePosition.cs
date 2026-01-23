using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Image;

/// <summary>
///     Position information for floating images.
/// </summary>
public record WordImagePosition
{
    /// <summary>
    ///     X position.
    /// </summary>
    [JsonPropertyName("x")]
    public required double X { get; init; }

    /// <summary>
    ///     Y position.
    /// </summary>
    [JsonPropertyName("y")]
    public required double Y { get; init; }

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
    ///     Wrap type.
    /// </summary>
    [JsonPropertyName("wrapType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? WrapType { get; init; }
}
