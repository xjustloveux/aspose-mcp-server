using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Page;

/// <summary>
///     Represents a PDF page box (media box, crop box, etc.).
/// </summary>
public record PdfPageBox
{
    /// <summary>
    ///     Lower-left X coordinate.
    /// </summary>
    [JsonPropertyName("llx")]
    public required double Llx { get; init; }

    /// <summary>
    ///     Lower-left Y coordinate.
    /// </summary>
    [JsonPropertyName("lly")]
    public required double Lly { get; init; }

    /// <summary>
    ///     Upper-right X coordinate.
    /// </summary>
    [JsonPropertyName("urx")]
    public required double Urx { get; init; }

    /// <summary>
    ///     Upper-right Y coordinate.
    /// </summary>
    [JsonPropertyName("ury")]
    public required double Ury { get; init; }
}
