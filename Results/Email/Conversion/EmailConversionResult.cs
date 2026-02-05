using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Conversion;

/// <summary>
///     Result containing email format conversion information.
/// </summary>
public record EmailConversionResult
{
    /// <summary>
    ///     Source file path of the email.
    /// </summary>
    [JsonPropertyName("sourcePath")]
    public required string SourcePath { get; init; }

    /// <summary>
    ///     Output file path of the converted email.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }

    /// <summary>
    ///     Source format of the email (e.g., "EML", "MSG").
    /// </summary>
    [JsonPropertyName("sourceFormat")]
    public required string SourceFormat { get; init; }

    /// <summary>
    ///     Target format of the converted email (e.g., "MSG", "HTML").
    /// </summary>
    [JsonPropertyName("targetFormat")]
    public required string TargetFormat { get; init; }

    /// <summary>
    ///     Output file size in bytes.
    /// </summary>
    [JsonPropertyName("fileSize")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? FileSize { get; init; }

    /// <summary>
    ///     Human-readable message describing the conversion result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
