using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Image;

/// <summary>
///     Information about a single image in Word document.
/// </summary>
public record WordImageInfo
{
    /// <summary>
    ///     Zero-based index of the image.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Image name.
    /// </summary>
    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; init; }

    /// <summary>
    ///     Image width.
    /// </summary>
    [JsonPropertyName("width")]
    public required double Width { get; init; }

    /// <summary>
    ///     Image height.
    /// </summary>
    [JsonPropertyName("height")]
    public required double Height { get; init; }

    /// <summary>
    ///     Whether the image is inline.
    /// </summary>
    [JsonPropertyName("isInline")]
    public required bool IsInline { get; init; }

    /// <summary>
    ///     Alignment for inline images.
    /// </summary>
    [JsonPropertyName("alignment")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Alignment { get; init; }

    /// <summary>
    ///     Position information for floating images.
    /// </summary>
    [JsonPropertyName("position")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WordImagePosition? Position { get; init; }

    /// <summary>
    ///     Context text around the image.
    /// </summary>
    [JsonPropertyName("context")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Context { get; init; }

    /// <summary>
    ///     Image type.
    /// </summary>
    [JsonPropertyName("imageType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ImageType { get; init; }

    /// <summary>
    ///     Original image size.
    /// </summary>
    [JsonPropertyName("originalSize")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ImageSize? OriginalSize { get; init; }

    /// <summary>
    ///     Hyperlink URL if any.
    /// </summary>
    [JsonPropertyName("hyperlink")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Hyperlink { get; init; }

    /// <summary>
    ///     Alternative text.
    /// </summary>
    [JsonPropertyName("altText")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AltText { get; init; }

    /// <summary>
    ///     Image title.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }
}
