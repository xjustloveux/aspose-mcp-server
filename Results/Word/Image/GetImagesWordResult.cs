using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Image;

/// <summary>
///     Result for getting images from Word documents.
/// </summary>
public record GetImagesWordResult
{
    /// <summary>
    ///     Number of images found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Section index if specified.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SectionIndex { get; init; }

    /// <summary>
    ///     List of image information.
    /// </summary>
    [JsonPropertyName("images")]
    public required IReadOnlyList<WordImageInfo> Images { get; init; }

    /// <summary>
    ///     Optional message when no images found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
