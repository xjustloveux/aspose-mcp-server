using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.FileOperations;

/// <summary>
///     Result containing detected email file format information.
/// </summary>
public record DetectFormatEmailResult
{
    /// <summary>
    ///     The detected format name (e.g., "Eml", "Msg").
    /// </summary>
    [JsonPropertyName("format")]
    public required string Format { get; init; }

    /// <summary>
    ///     The expected file extension for the detected format.
    /// </summary>
    [JsonPropertyName("extension")]
    public required string Extension { get; init; }

    /// <summary>
    ///     Human-readable message describing the detection result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
