using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Signature;

/// <summary>
///     Result for getting signatures from PDF documents.
/// </summary>
public record GetSignaturesResult
{
    /// <summary>
    ///     Number of signatures.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of signature information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<SignatureInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no signatures found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
