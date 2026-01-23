using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Table;

/// <summary>
///     Result for getting table structure from Word documents.
/// </summary>
public sealed record GetTableStructureWordResult
{
    /// <summary>
    ///     The formatted table structure information as plain text.
    /// </summary>
    [JsonPropertyName("content")]
    public required string Content { get; init; }
}
