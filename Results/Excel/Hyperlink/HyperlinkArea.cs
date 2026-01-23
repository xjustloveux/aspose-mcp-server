using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Hyperlink;

/// <summary>
///     Cell area for a hyperlink.
/// </summary>
public record HyperlinkArea
{
    /// <summary>
    ///     Start cell reference.
    /// </summary>
    [JsonPropertyName("startCell")]
    public required string StartCell { get; init; }

    /// <summary>
    ///     End cell reference.
    /// </summary>
    [JsonPropertyName("endCell")]
    public required string EndCell { get; init; }
}
