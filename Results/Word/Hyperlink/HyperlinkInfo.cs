using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Hyperlink;

/// <summary>
///     Information about a single hyperlink.
/// </summary>
public record HyperlinkInfo
{
    /// <summary>
    ///     Zero-based index of the hyperlink.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Display text of the hyperlink.
    /// </summary>
    [JsonPropertyName("displayText")]
    public required string DisplayText { get; init; }

    /// <summary>
    ///     URL or address of the hyperlink.
    /// </summary>
    [JsonPropertyName("address")]
    public required string Address { get; init; }

    /// <summary>
    ///     Sub-address (e.g., bookmark within document).
    /// </summary>
    [JsonPropertyName("subAddress")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SubAddress { get; init; }

    /// <summary>
    ///     Tooltip text for the hyperlink.
    /// </summary>
    [JsonPropertyName("tooltip")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Tooltip { get; init; }

    /// <summary>
    ///     Index of the paragraph containing the hyperlink.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ParagraphIndex { get; init; }
}
