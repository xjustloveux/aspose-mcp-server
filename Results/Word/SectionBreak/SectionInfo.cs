using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Information about a document section.
/// </summary>
public record SectionInfo
{
    /// <summary>
    ///     The index of the section.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Section break information.
    /// </summary>
    [JsonPropertyName("sectionBreak")]
    public required SectionBreakInfo SectionBreak { get; init; }

    /// <summary>
    ///     Page setup information.
    /// </summary>
    [JsonPropertyName("pageSetup")]
    public required PageSetupInfo PageSetup { get; init; }

    /// <summary>
    ///     Content statistics for the section.
    /// </summary>
    [JsonPropertyName("contentStatistics")]
    public required ContentStatisticsInfo ContentStatistics { get; init; }

    /// <summary>
    ///     Headers and footers count.
    /// </summary>
    [JsonPropertyName("headersFooters")]
    public required HeadersFootersCountInfo HeadersFooters { get; init; }
}
