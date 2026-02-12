using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Result for send command operation.
/// </summary>
public record SendCommandResult
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
    public string SessionId { get; init; } = string.Empty;

    /// <summary>
    ///     Extension ID.
    /// </summary>
    [JsonPropertyName("extensionId")]
    public string ExtensionId { get; init; } = string.Empty;

    /// <summary>
    ///     Command ID for tracking.
    /// </summary>
    [JsonPropertyName("commandId")]
    public string? CommandId { get; init; }

    /// <summary>
    ///     Command type that was sent.
    /// </summary>
    [JsonPropertyName("commandType")]
    public string CommandType { get; init; } = string.Empty;

    /// <summary>
    ///     Result data from the extension.
    /// </summary>
    [JsonPropertyName("result")]
    public Dictionary<string, object>? Result { get; init; }
}
