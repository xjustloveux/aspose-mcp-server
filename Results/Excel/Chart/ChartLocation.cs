using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Chart;

/// <summary>
///     Chart location information.
/// </summary>
public record ChartLocation
{
    /// <summary>
    ///     Upper left row index.
    /// </summary>
    [JsonPropertyName("upperLeftRow")]
    public required int UpperLeftRow { get; init; }

    /// <summary>
    ///     Lower right row index.
    /// </summary>
    [JsonPropertyName("lowerRightRow")]
    public required int LowerRightRow { get; init; }

    /// <summary>
    ///     Upper left column index.
    /// </summary>
    [JsonPropertyName("upperLeftColumn")]
    public required int UpperLeftColumn { get; init; }

    /// <summary>
    ///     Lower right column index.
    /// </summary>
    [JsonPropertyName("lowerRightColumn")]
    public required int LowerRightColumn { get; init; }
}
