using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.BarCode.Recognize;

/// <summary>
///     Information about a single recognized barcode.
/// </summary>
public record BarcodeInfo
{
    /// <summary>
    ///     The decoded text content of the barcode.
    /// </summary>
    [JsonPropertyName("codeText")]
    public required string CodeText { get; init; }

    /// <summary>
    ///     The type/symbology of the recognized barcode (e.g., "QR", "Code128").
    /// </summary>
    [JsonPropertyName("codeType")]
    public required string CodeType { get; init; }

    /// <summary>
    ///     The confidence level of the recognition (if available).
    /// </summary>
    [JsonPropertyName("confidence")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Confidence { get; init; }

    /// <summary>
    ///     The region of the barcode in the image (x, y, width, height).
    /// </summary>
    [JsonPropertyName("region")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Region { get; init; }
}
