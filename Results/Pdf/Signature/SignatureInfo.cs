using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Signature;

/// <summary>
///     Information about a single signature.
/// </summary>
public record SignatureInfo
{
    /// <summary>
    ///     Zero-based index of the signature.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Signature name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Whether signature is valid.
    /// </summary>
    [JsonPropertyName("isValid")]
    public required bool IsValid { get; init; }

    /// <summary>
    ///     Whether signature has certificate.
    /// </summary>
    [JsonPropertyName("hasCertificate")]
    public required bool HasCertificate { get; init; }
}
