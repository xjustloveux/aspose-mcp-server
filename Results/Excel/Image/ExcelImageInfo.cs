using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Image;

/// <summary>
///     Information about a single Excel image.
/// </summary>
public record ExcelImageInfo
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
    ///     Alternative text.
    /// </summary>
    [JsonPropertyName("alternativeText")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AlternativeText { get; init; }

    /// <summary>
    ///     Image type.
    /// </summary>
    [JsonPropertyName("imageType")]
    public required string ImageType { get; init; }

    /// <summary>
    ///     Image location in worksheet.
    /// </summary>
    [JsonPropertyName("location")]
    public required ExcelImageLocation Location { get; init; }

    /// <summary>
    ///     Image width.
    /// </summary>
    [JsonPropertyName("width")]
    public required int Width { get; init; }

    /// <summary>
    ///     Image height.
    /// </summary>
    [JsonPropertyName("height")]
    public required int Height { get; init; }

    /// <summary>
    ///     Whether aspect ratio is locked.
    /// </summary>
    [JsonPropertyName("isLockAspectRatio")]
    public required bool IsLockAspectRatio { get; init; }
}
