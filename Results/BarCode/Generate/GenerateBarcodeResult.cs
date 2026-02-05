using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.BarCode.Generate;

/// <summary>
///     Result containing barcode generation information.
/// </summary>
public record GenerateBarcodeResult
{
    /// <summary>
    ///     Output file path of the generated barcode image.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }

    /// <summary>
    ///     The barcode type used for generation (e.g., "QR", "Code128").
    /// </summary>
    [JsonPropertyName("barcodeType")]
    public required string BarcodeType { get; init; }

    /// <summary>
    ///     The text encoded in the barcode.
    /// </summary>
    [JsonPropertyName("encodedText")]
    public required string EncodedText { get; init; }

    /// <summary>
    ///     The image format of the output file (e.g., "PNG", "JPEG").
    /// </summary>
    [JsonPropertyName("imageFormat")]
    public required string ImageFormat { get; init; }

    /// <summary>
    ///     Output file size in bytes.
    /// </summary>
    [JsonPropertyName("fileSize")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? FileSize { get; init; }

    /// <summary>
    ///     Human-readable message describing the generation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
