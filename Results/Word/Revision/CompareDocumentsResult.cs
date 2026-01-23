using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Revision;

/// <summary>
///     Result for comparing two Word documents.
/// </summary>
public record CompareDocumentsResult
{
    /// <summary>
    ///     Number of differences (revisions) found between the documents.
    /// </summary>
    [JsonPropertyName("revisionCount")]
    public required int RevisionCount { get; init; }

    /// <summary>
    ///     Path to the output comparison document.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }
}
