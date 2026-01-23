using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Section;

/// <summary>
///     Result for getting sections from PowerPoint presentations.
/// </summary>
public record GetSectionsResult
{
    /// <summary>
    ///     Number of sections.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of section information.
    /// </summary>
    [JsonPropertyName("sections")]
    public required IReadOnlyList<SectionInfo> Sections { get; init; }

    /// <summary>
    ///     Optional message when no sections found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
