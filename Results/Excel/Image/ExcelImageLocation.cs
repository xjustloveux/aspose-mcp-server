using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Image;

/// <summary>
///     Location information for an Excel image.
/// </summary>
public record ExcelImageLocation
{
    /// <summary>
    ///     Upper left cell reference.
    /// </summary>
    [JsonPropertyName("upperLeftCell")]
    public required string UpperLeftCell { get; init; }

    /// <summary>
    ///     Lower right cell reference.
    /// </summary>
    [JsonPropertyName("lowerRightCell")]
    public required string LowerRightCell { get; init; }

    /// <summary>
    ///     Upper left row index.
    /// </summary>
    [JsonPropertyName("upperLeftRow")]
    public required int UpperLeftRow { get; init; }

    /// <summary>
    ///     Upper left column index.
    /// </summary>
    [JsonPropertyName("upperLeftColumn")]
    public required int UpperLeftColumn { get; init; }

    /// <summary>
    ///     Lower right row index.
    /// </summary>
    [JsonPropertyName("lowerRightRow")]
    public required int LowerRightRow { get; init; }

    /// <summary>
    ///     Lower right column index.
    /// </summary>
    [JsonPropertyName("lowerRightColumn")]
    public required int LowerRightColumn { get; init; }
}
