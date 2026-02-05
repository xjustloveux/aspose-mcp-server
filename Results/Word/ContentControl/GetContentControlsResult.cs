using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.ContentControl;

/// <summary>
///     Result for getting content controls from Word documents.
/// </summary>
public record GetContentControlsResult
{
    /// <summary>
    ///     Total number of content controls in the document.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of content control information.
    /// </summary>
    [JsonPropertyName("contentControls")]
    public required IReadOnlyList<ContentControlInfo> ContentControls { get; init; }

    /// <summary>
    ///     Optional message when no content controls found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
