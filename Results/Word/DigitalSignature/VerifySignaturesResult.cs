using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.DigitalSignature;

/// <summary>
///     Result for verifying digital signatures in a Word document.
/// </summary>
public record VerifySignaturesResult
{
    /// <summary>
    ///     Human-readable message describing the verification result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }

    /// <summary>
    ///     Whether all signatures are valid.
    /// </summary>
    [JsonPropertyName("allValid")]
    public bool AllValid { get; init; }

    /// <summary>
    ///     The total number of digital signatures.
    /// </summary>
    [JsonPropertyName("totalCount")]
    public int TotalCount { get; init; }

    /// <summary>
    ///     The number of valid signatures.
    /// </summary>
    [JsonPropertyName("validCount")]
    public int ValidCount { get; init; }
}
