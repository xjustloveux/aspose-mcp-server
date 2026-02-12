using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Result for bind operation.
/// </summary>
public record BindExtensionResult
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
    ///     Binding information if successful.
    /// </summary>
    [JsonPropertyName("binding")]
    public BindingInfoDto? Binding { get; init; }
}
