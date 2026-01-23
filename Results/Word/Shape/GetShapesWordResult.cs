using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Shape;

/// <summary>
///     Result for getting shapes from Word documents.
/// </summary>
public sealed record GetShapesWordResult
{
    /// <summary>
    ///     The formatted shape information as plain text.
    /// </summary>
    [JsonPropertyName("content")]
    public required string Content { get; init; }
}
