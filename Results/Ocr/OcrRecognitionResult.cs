using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Ocr;

/// <summary>
///     Complete OCR recognition result containing all pages and overall statistics.
/// </summary>
public record OcrRecognitionResult
{
    /// <summary>
    ///     Full recognized text content across all pages.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Overall recognition confidence score (0.0 to 1.0).
    /// </summary>
    [JsonPropertyName("confidence")]
    public double Confidence { get; init; }

    /// <summary>
    ///     Number of pages processed.
    /// </summary>
    [JsonPropertyName("pageCount")]
    public int PageCount { get; init; }

    /// <summary>
    ///     Per-page recognition results.
    /// </summary>
    [JsonPropertyName("pages")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<OcrPageResult>? Pages { get; init; }
}
