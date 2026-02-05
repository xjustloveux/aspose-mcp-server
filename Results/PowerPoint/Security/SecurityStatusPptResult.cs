using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Security;

/// <summary>
///     Result containing the security status of a presentation.
/// </summary>
public record SecurityStatusPptResult
{
    /// <summary>
    ///     Whether the presentation is encrypted.
    /// </summary>
    [JsonPropertyName("isEncrypted")]
    public required bool IsEncrypted { get; init; }

    /// <summary>
    ///     Whether the presentation has write protection.
    /// </summary>
    [JsonPropertyName("isWriteProtected")]
    public required bool IsWriteProtected { get; init; }

    /// <summary>
    ///     Whether the presentation is marked as final.
    /// </summary>
    [JsonPropertyName("isMarkedFinal")]
    public required bool IsMarkedFinal { get; init; }

    /// <summary>
    ///     Whether the presentation is read-only recommended.
    /// </summary>
    [JsonPropertyName("isReadOnlyRecommended")]
    public required bool IsReadOnlyRecommended { get; init; }

    /// <summary>
    ///     Human-readable message describing the security status.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
