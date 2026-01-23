using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Text;

/// <summary>
///     Information about a single text fragment including font details.
/// </summary>
public record PdfTextFragment
{
    /// <summary>
    ///     Text content of the fragment.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Font name used for the text.
    /// </summary>
    [JsonPropertyName("fontName")]
    public required string FontName { get; init; }

    /// <summary>
    ///     Font size used for the text.
    /// </summary>
    [JsonPropertyName("fontSize")]
    public required double FontSize { get; init; }
}
