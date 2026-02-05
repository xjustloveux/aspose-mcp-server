using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.BarCode.Recognize;

/// <summary>
///     Result containing barcode recognition information.
/// </summary>
public record RecognizeBarcodeResult
{
    /// <summary>
    ///     Source image file path.
    /// </summary>
    [JsonPropertyName("sourcePath")]
    public required string SourcePath { get; init; }

    /// <summary>
    ///     List of recognized barcodes.
    /// </summary>
    [JsonPropertyName("barcodes")]
    public required List<BarcodeInfo> Barcodes { get; init; }

    /// <summary>
    ///     Total number of barcodes recognized.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     The decode type filter used for recognition.
    /// </summary>
    [JsonPropertyName("decodeType")]
    public required string DecodeType { get; init; }

    /// <summary>
    ///     Human-readable message describing the recognition result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
