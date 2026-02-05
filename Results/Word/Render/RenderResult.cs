using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Render;

/// <summary>
///     Result for rendering a Word document page to an image.
/// </summary>
public record RenderResult
{
    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }

    /// <summary>
    ///     The output image file path(s).
    /// </summary>
    [JsonPropertyName("outputPaths")]
    public required List<string> OutputPaths { get; init; }

    /// <summary>
    ///     The image format used for rendering.
    /// </summary>
    [JsonPropertyName("format")]
    public required string Format { get; init; }
}
