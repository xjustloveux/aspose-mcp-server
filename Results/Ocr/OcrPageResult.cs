using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Ocr;

/// <summary>
///     OCR recognition result for a single page.
/// </summary>
public record OcrPageResult
{
    /// <summary>
    ///     Zero-based page index.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    public int PageIndex { get; init; }

    /// <summary>
    ///     Full recognized text content of the page.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Overall recognition confidence score for the page (0.0 to 1.0).
    /// </summary>
    [JsonPropertyName("confidence")]
    public double Confidence { get; init; }

    /// <summary>
    ///     List of recognized words with position and confidence details.
    /// </summary>
    [JsonPropertyName("words")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<OcrWordInfo>? Words { get; init; }
}
