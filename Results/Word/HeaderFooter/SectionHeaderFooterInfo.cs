using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.HeaderFooter;

/// <summary>
///     Header and footer information for a section.
/// </summary>
public record SectionHeaderFooterInfo
{
    /// <summary>
    ///     Section index.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    public required int SectionIndex { get; init; }

    /// <summary>
    ///     Headers content.
    /// </summary>
    [JsonPropertyName("headers")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyDictionary<string, string?>? Headers { get; init; }

    /// <summary>
    ///     Footers content.
    /// </summary>
    [JsonPropertyName("footers")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyDictionary<string, string?>? Footers { get; init; }
}
