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
    ///     Recognition confidence score (0.0 to 1.0).
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
