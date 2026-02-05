using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.FileOperations;

/// <summary>
///     Result containing email format conversion details.
/// </summary>
public record EmailConversionResult
{
    /// <summary>
    ///     The source email file path.
    /// </summary>
    [JsonPropertyName("sourcePath")]
    public required string SourcePath { get; init; }

    /// <summary>
    ///     The output file path after conversion.
    /// </summary>
    [JsonPropertyName("outputPath")]
    public required string OutputPath { get; init; }

    /// <summary>
    ///     The source email format.
    /// </summary>
    [JsonPropertyName("sourceFormat")]
    public required string SourceFormat { get; init; }

    /// <summary>
    ///     The target email format.
    /// </summary>
    [JsonPropertyName("targetFormat")]
    public required string TargetFormat { get; init; }

    /// <summary>
    ///     Human-readable message describing the conversion result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
