using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Result for list extensions operation.
/// </summary>
public record ListExtensionsResult
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
    ///     Number of extensions.
    /// </summary>
    [JsonPropertyName("count")]
    public int Count { get; init; }

    /// <summary>
    ///     List of available extensions.
    /// </summary>
    [JsonPropertyName("extensions")]
    public required IReadOnlyList<ExtensionInfoDto> Extensions { get; init; }
}
