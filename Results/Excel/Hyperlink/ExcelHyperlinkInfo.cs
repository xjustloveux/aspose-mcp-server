using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Hyperlink;

/// <summary>
///     Information about a single Excel hyperlink.
/// </summary>
public record ExcelHyperlinkInfo
{
    /// <summary>
    ///     Zero-based index of the hyperlink.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Cell reference.
    /// </summary>
    [JsonPropertyName("cell")]
    public required string Cell { get; init; }

    /// <summary>
    ///     Hyperlink URL.
    /// </summary>
    [JsonPropertyName("url")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Url { get; init; }

    /// <summary>
    ///     Display text.
    /// </summary>
    [JsonPropertyName("displayText")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DisplayText { get; init; }

    /// <summary>
    ///     Cell area covered by the hyperlink.
    /// </summary>
    [JsonPropertyName("area")]
    public required HyperlinkArea Area { get; init; }
}
