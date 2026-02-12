using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Result for status operation.
/// </summary>
public record ExtensionStatusResult
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
    ///     Extension ID.
    /// </summary>
    [JsonPropertyName("extensionId")]
    public required string ExtensionId { get; init; }

    /// <summary>
    ///     Extension name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Whether the extension is available.
    /// </summary>
    [JsonPropertyName("isAvailable")]
    public bool IsAvailable { get; init; }

    /// <summary>
    ///     Whether the extension is currently initializing.
    /// </summary>
    [JsonPropertyName("isInitializing")]
    public bool IsInitializing { get; init; }

    /// <summary>
    ///     Reason if unavailable.
    /// </summary>
    [JsonPropertyName("unavailableReason")]
    public string? UnavailableReason { get; init; }

    /// <summary>
    ///     Current state.
    /// </summary>
    [JsonPropertyName("state")]
    public required string State { get; init; }

    /// <summary>
    ///     Last activity time.
    /// </summary>
    [JsonPropertyName("lastActivity")]
    public DateTime? LastActivity { get; init; }

    /// <summary>
    ///     Number of restart attempts.
    /// </summary>
    [JsonPropertyName("restartCount")]
    public int RestartCount { get; init; }

    /// <summary>
    ///     Number of active bindings.
    /// </summary>
    [JsonPropertyName("activeBindings")]
    public int ActiveBindings { get; init; }
}
