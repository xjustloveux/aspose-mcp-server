using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Revision;

/// <summary>
///     Information about a single revision.
/// </summary>
public sealed record RevisionInfo
{
    /// <summary>
    ///     Zero-based index of the revision.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Type of revision (Insertion, Deletion, FormatChange, etc.).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Author of the revision.
    /// </summary>
    [JsonPropertyName("author")]
    public required string Author { get; init; }

    /// <summary>
    ///     Date of the revision formatted as yyyy-MM-dd HH:mm:ss.
    /// </summary>
    [JsonPropertyName("date")]
    public required string Date { get; init; }

    /// <summary>
    ///     Text preview of the revision (truncated to 100 characters).
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }
}
