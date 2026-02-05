using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.DigitalSignature;

/// <summary>
///     Result for listing digital signatures in a Word document.
/// </summary>
public record GetSignaturesResult
{
    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }

    /// <summary>
    ///     The total number of digital signatures found.
    /// </summary>
    [JsonPropertyName("count")]
    public int Count { get; init; }

    /// <summary>
    ///     The list of digital signature information.
    /// </summary>
    [JsonPropertyName("signatures")]
    public required List<SignatureInfo> Signatures { get; init; }
}
