using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.DigitalSignature;

/// <summary>
///     Information about a digital signature in a Word document.
/// </summary>
public record SignatureInfo
{
    /// <summary>
    ///     The signer's name from the certificate.
    /// </summary>
    [JsonPropertyName("signerName")]
    public string? SignerName { get; init; }

    /// <summary>
    ///     The comments associated with the signature.
    /// </summary>
    [JsonPropertyName("comments")]
    public string? Comments { get; init; }

    /// <summary>
    ///     The date and time the signature was applied.
    /// </summary>
    [JsonPropertyName("signTime")]
    public string? SignTime { get; init; }

    /// <summary>
    ///     Whether the signature is valid.
    /// </summary>
    [JsonPropertyName("isValid")]
    public bool IsValid { get; init; }

    /// <summary>
    ///     The certificate issuer name.
    /// </summary>
    [JsonPropertyName("issuerName")]
    public string? IssuerName { get; init; }

    /// <summary>
    ///     The certificate subject name.
    /// </summary>
    [JsonPropertyName("subjectName")]
    public string? SubjectName { get; init; }
}
