using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Result for getting sections from Word documents.
/// </summary>
public record GetSectionsWordResult
{
    /// <summary>
    ///     Total number of sections.
    /// </summary>
    [JsonPropertyName("totalSections")]
    public required int TotalSections { get; init; }

    /// <summary>
    ///     Single section info (when querying specific section).
    /// </summary>
    [JsonPropertyName("section")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public SectionInfo? Section { get; init; }

    /// <summary>
    ///     List of section information (when querying all sections).
    /// </summary>
    [JsonPropertyName("sections")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<SectionInfo>? Sections { get; init; }
}
