using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Revision;

/// <summary>
///     Result for getting revisions from Word documents.
/// </summary>
public sealed record GetRevisionsWordResult
{
    /// <summary>
    ///     Total number of revisions.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of revisions.
    /// </summary>
    [JsonPropertyName("revisions")]
    public required IReadOnlyList<RevisionInfo> Revisions { get; init; }
}
