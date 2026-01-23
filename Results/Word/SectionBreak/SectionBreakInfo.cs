using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Section break type information.
/// </summary>
public record SectionBreakInfo
{
    /// <summary>
    ///     The type of section break.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }
}
