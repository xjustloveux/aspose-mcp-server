using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Link;

/// <summary>
///     Result for getting links from PDF documents.
/// </summary>
public record GetLinksResult
{
    /// <summary>
    ///     Total number of links found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Page index if filtering by specific page.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PageIndex { get; init; }

    /// <summary>
    ///     List of link information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<LinkInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no links found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
