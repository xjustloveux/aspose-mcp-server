using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Shape;

/// <summary>
///     Result for getting textboxes from Word documents.
/// </summary>
public sealed record GetTextboxesWordResult
{
    /// <summary>
    ///     The formatted textbox information as plain text.
    /// </summary>
    [JsonPropertyName("content")]
    public required string Content { get; init; }
}
