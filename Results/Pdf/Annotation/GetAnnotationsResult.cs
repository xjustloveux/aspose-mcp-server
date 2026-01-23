using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Annotation;

/// <summary>
///     Result for getting annotations from PDF documents.
/// </summary>
public record GetAnnotationsResult
{
    /// <summary>
    ///     Number of annotations.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of annotation information.
    /// </summary>
    [JsonPropertyName("annotations")]
    public required IReadOnlyList<AnnotationInfo> Annotations { get; init; }
}
