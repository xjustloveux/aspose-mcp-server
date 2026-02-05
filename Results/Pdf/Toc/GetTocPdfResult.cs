using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Toc;

/// <summary>
///     Result for getting table of contents from PDF documents.
/// </summary>
public record GetTocPdfResult
{
    /// <summary>
    ///     Whether the document has a table of contents.
    /// </summary>
    [JsonPropertyName("hasToc")]
    public required bool HasToc { get; init; }

    /// <summary>
    ///     Number of TOC entries found.
    /// </summary>
    [JsonPropertyName("entryCount")]
    public required int EntryCount { get; init; }

    /// <summary>
    ///     List of TOC entry information.
    /// </summary>
    [JsonPropertyName("entries")]
    public required IReadOnlyList<TocEntryPdfInfo> Entries { get; init; }

    /// <summary>
    ///     Human-readable message describing the result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
