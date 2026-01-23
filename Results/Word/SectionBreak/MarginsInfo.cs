using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Page margin values.
/// </summary>
public record MarginsInfo
{
    /// <summary>
    ///     Top margin in points.
    /// </summary>
    [JsonPropertyName("top")]
    public required double Top { get; init; }

    /// <summary>
    ///     Bottom margin in points.
    /// </summary>
    [JsonPropertyName("bottom")]
    public required double Bottom { get; init; }

    /// <summary>
    ///     Left margin in points.
    /// </summary>
    [JsonPropertyName("left")]
    public required double Left { get; init; }

    /// <summary>
    ///     Right margin in points.
    /// </summary>
    [JsonPropertyName("right")]
    public required double Right { get; init; }
}
