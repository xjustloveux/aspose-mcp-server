using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Ocr;

/// <summary>
///     Result for OCR-based PDF to editable document conversion operations.
/// </summary>
public record OcrConversionResult
{
    /// <summary>
    ///     Source PDF file path.
    /// </summary>
    [JsonPropertyName("sourcePath")]
    public required string SourcePath { get; init; }

    /// <summary>
    ///     Output file path of the converted document.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }

    /// <summary>
    ///     Target format of the converted document (e.g., "docx", "xlsx", "pptx").
    /// </summary>
    [JsonPropertyName("targetFormat")]
    public required string TargetFormat { get; init; }

    /// <summary>
    ///     Number of pages processed.
    /// </summary>
    [JsonPropertyName("pageCount")]
    public int PageCount { get; init; }

    /// <summary>
    ///     Average recognition confidence score across all pages (0.0 to 1.0).
    /// </summary>
    [JsonPropertyName("averageConfidence")]
    public double AverageConfidence { get; init; }

    /// <summary>
    ///     Output file size in bytes.
    /// </summary>
    [JsonPropertyName("fileSize")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? FileSize { get; init; }

    /// <summary>
    ///     Human-readable message describing the conversion result.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
