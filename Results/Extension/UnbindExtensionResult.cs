using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Result for unbind operation.
/// </summary>
public record UnbindExtensionResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Error code if failed.
    /// </summary>
    [JsonPropertyName("errorCode")]
    public ExtensionErrorCode ErrorCode { get; init; } = ExtensionErrorCode.None;

    /// <summary>
    ///     Error message if failed.
    /// </summary>
    [JsonPropertyName("error")]
    public string? Error { get; init; }

    /// <summary>
    ///     Session ID that was unbound.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Extension ID that was unbound (null if all extensions were unbound).
    /// </summary>
    [JsonPropertyName("extensionId")]
    public string? ExtensionId { get; init; }

    /// <summary>
    ///     Number of bindings removed.
    /// </summary>
    [JsonPropertyName("unboundCount")]
    public int UnboundCount { get; init; }
}
