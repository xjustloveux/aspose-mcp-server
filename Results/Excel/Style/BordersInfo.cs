using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Style;

/// <summary>
///     Border information for all sides.
/// </summary>
public record BordersInfo
{
    /// <summary>
    ///     Top border information.
    /// </summary>
    [JsonPropertyName("top")]
    public required BorderInfo Top { get; init; }

    /// <summary>
    ///     Bottom border information.
    /// </summary>
    [JsonPropertyName("bottom")]
    public required BorderInfo Bottom { get; init; }

    /// <summary>
    ///     Left border information.
    /// </summary>
    [JsonPropertyName("left")]
    public required BorderInfo Left { get; init; }

    /// <summary>
    ///     Right border information.
    /// </summary>
    [JsonPropertyName("right")]
    public required BorderInfo Right { get; init; }
}
