using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Headers and footers count information.
/// </summary>
public record HeadersFootersCountInfo
{
    /// <summary>
    ///     Number of headers with content.
    /// </summary>
    [JsonPropertyName("headers")]
    public required int Headers { get; init; }

    /// <summary>
    ///     Number of footers with content.
    /// </summary>
    [JsonPropertyName("footers")]
    public required int Footers { get; init; }
}
