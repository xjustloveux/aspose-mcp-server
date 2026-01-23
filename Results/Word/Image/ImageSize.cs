using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Image;

/// <summary>
///     Image size in pixels.
/// </summary>
public record ImageSize
{
    /// <summary>
    ///     Width in pixels.
    /// </summary>
    [JsonPropertyName("widthPixels")]
    public required int WidthPixels { get; init; }

    /// <summary>
    ///     Height in pixels.
    /// </summary>
    [JsonPropertyName("heightPixels")]
    public required int HeightPixels { get; init; }
}
