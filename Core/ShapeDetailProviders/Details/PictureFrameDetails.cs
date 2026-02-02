using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for PictureFrame elements.
/// </summary>
public sealed record PictureFrameDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the hyperlink target, if any.
    /// </summary>
    [JsonPropertyName("hyperlink")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Hyperlink { get; init; }

    /// <summary>
    ///     Gets the content type of the image.
    /// </summary>
    [JsonPropertyName("contentType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContentType { get; init; }

    /// <summary>
    ///     Gets the image width in pixels.
    /// </summary>
    [JsonPropertyName("imageWidth")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ImageWidth { get; init; }

    /// <summary>
    ///     Gets the image height in pixels.
    /// </summary>
    [JsonPropertyName("imageHeight")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ImageHeight { get; init; }

    /// <summary>
    ///     Gets the left crop value.
    /// </summary>
    [JsonPropertyName("cropLeft")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public float? CropLeft { get; init; }

    /// <summary>
    ///     Gets the right crop value.
    /// </summary>
    [JsonPropertyName("cropRight")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public float? CropRight { get; init; }

    /// <summary>
    ///     Gets the top crop value.
    /// </summary>
    [JsonPropertyName("cropTop")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public float? CropTop { get; init; }

    /// <summary>
    ///     Gets the bottom crop value.
    /// </summary>
    [JsonPropertyName("cropBottom")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public float? CropBottom { get; init; }
}
