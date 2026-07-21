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
    ///     Recognition confidence score for the page. The bundled Aspose.OCR engine does not
    ///     report confidence values, so this is always 0; it is kept for wire-format stability.
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
