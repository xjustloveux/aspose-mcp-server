using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Result for set_format operation.
/// </summary>
public record SetFormatResult
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
    ///     Session ID.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Extension ID.
    /// </summary>
    [JsonPropertyName("extensionId")]
    public required string ExtensionId { get; init; }

    /// <summary>
    ///     New output format.
    /// </summary>
    [JsonPropertyName("newFormat")]
    public required string NewFormat { get; init; }
}
