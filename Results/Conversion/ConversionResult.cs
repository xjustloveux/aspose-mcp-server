using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Conversion;

/// <summary>
///     Result for file conversion operations.
/// </summary>
public record ConversionResult
{
    /// <summary>
    ///     Source file path.
    /// </summary>
    [JsonPropertyName("sourcePath")]
    public required string SourcePath { get; init; }

    /// <summary>
    ///     Output file path.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }

    /// <summary>
    ///     Source format (detected or specified).
    /// </summary>
    [JsonPropertyName("sourceFormat")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceFormat { get; init; }

    /// <summary>
    ///     Target format.
    /// </summary>
    [JsonPropertyName("targetFormat")]
    public required string TargetFormat { get; init; }

    /// <summary>
    ///     File size in bytes.
    /// </summary>
    [JsonPropertyName("fileSize")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? FileSize { get; init; }

    /// <summary>
    ///     Human-readable message.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
