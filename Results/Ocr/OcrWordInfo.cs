using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Ocr;

/// <summary>
///     Information about a single recognized word including text, confidence, and position.
/// </summary>
public record OcrWordInfo
{
    /// <summary>
    ///     The recognized text content.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Recognition confidence score for the word. The bundled Aspose.OCR engine does not
    ///     report confidence values, so this is always 0; it is kept for wire-format stability.
    /// </summary>
    [JsonPropertyName("confidence")]
    public double Confidence { get; init; }

    /// <summary>
    ///     Bounding box coordinates for the recognized word.
    /// </summary>
    [JsonPropertyName("boundingBox")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public OcrBoundingBox? BoundingBox { get; init; }
}
