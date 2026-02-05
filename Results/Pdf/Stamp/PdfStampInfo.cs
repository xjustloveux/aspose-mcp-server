using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Stamp;

/// <summary>
///     Information about a single stamp annotation in a PDF document.
/// </summary>
public record PdfStampInfo
{
    /// <summary>
    ///     The 1-based page index where the stamp is located.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    public required int PageIndex { get; init; }

    /// <summary>
    ///     The 1-based annotation index on the page.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     The stamp type (e.g., "stamp").
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     The stamp annotation contents.
    /// </summary>
    [JsonPropertyName("contents")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Contents { get; init; }

    /// <summary>
    ///     The stamp opacity value between 0.0 and 1.0.
    /// </summary>
    [JsonPropertyName("opacity")]
    public required double Opacity { get; init; }
}
