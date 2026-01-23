using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Annotation;

/// <summary>
///     Information about a single annotation.
/// </summary>
public record AnnotationInfo
{
    /// <summary>
    ///     Page index.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    public required int PageIndex { get; init; }

    /// <summary>
    ///     Annotation index within the page.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Annotation type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Annotation title.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     Annotation subject.
    /// </summary>
    [JsonPropertyName("subject")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; init; }

    /// <summary>
    ///     Annotation contents.
    /// </summary>
    [JsonPropertyName("contents")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Contents { get; init; }

    /// <summary>
    ///     Annotation rectangle.
    /// </summary>
    [JsonPropertyName("rect")]
    public required AnnotationRect Rect { get; init; }
}
