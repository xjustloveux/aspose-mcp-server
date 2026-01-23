using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Layout;

/// <summary>
///     Result type for getting master slides from PowerPoint presentations.
/// </summary>
public sealed record GetMastersResult
{
    /// <summary>
    ///     Gets the count of masters.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Gets the masters list.
    /// </summary>
    [JsonPropertyName("masters")]
    public required List<GetMasterInfo> Masters { get; init; }

    /// <summary>
    ///     Gets the message (when no masters found).
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
