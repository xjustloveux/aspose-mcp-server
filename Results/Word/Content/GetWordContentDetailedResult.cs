using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Content;

/// <summary>
///     Result for getting detailed Word document content including headers and footers.
/// </summary>
public sealed record GetWordContentDetailedResult
{
    /// <summary>
    ///     The detailed document content as plain text.
    /// </summary>
    [JsonPropertyName("content")]
    public required string Content { get; init; }
}
