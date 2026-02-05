using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Ocr;

/// <summary>
///     Result for OCR image preprocessing operations.
/// </summary>
public record OcrPreprocessingResult
{
    /// <summary>
    ///     Source image file path.
    /// </summary>
    [JsonPropertyName("sourcePath")]
    public required string SourcePath { get; init; }

    /// <summary>
    ///     Output file path of the preprocessed image.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }

    /// <summary>
    ///     Preprocessing operation applied (e.g., "auto_skew", "denoise", "contrast").
    /// </summary>
    [JsonPropertyName("operation")]
    public required string Operation { get; init; }

    /// <summary>
    ///     Output file size in bytes.
    /// </summary>
    [JsonPropertyName("fileSize")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? FileSize { get; init; }

    /// <summary>
    ///     Human-readable message describing the preprocessing result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
