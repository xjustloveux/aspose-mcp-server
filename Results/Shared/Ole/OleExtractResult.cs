using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Shared.Ole;

/// <summary>
///     Result payload returned by the <c>extract</c> operation — extraction of a single
///     OLE object by index. Path-only return (no inline bytes) per spec §outputs.
/// </summary>
public sealed record OleExtractResult
{
    /// <summary>
    ///     Zero-based index of the extracted OLE object in the container snapshot at
    ///     call time.
    /// </summary>
    [JsonPropertyName("index")]
    [JsonPropertyOrder(1)]
    public int Index { get; init; }

    /// <summary>
    ///     Absolute path to the written file. Guaranteed to reside within the validated
    ///     <c>outputDirectory</c>.
    /// </summary>
    [JsonPropertyName("outputFilePath")]
    [JsonPropertyOrder(2)]
    public string OutputFilePath { get; init; } = string.Empty;

    /// <summary>
    ///     Byte count written to disk. Matches the on-disk file size.
    /// </summary>
    [JsonPropertyName("bytesWritten")]
    [JsonPropertyOrder(3)]
    public long BytesWritten { get; init; }

    /// <summary>
    ///     <c>true</c> when the suggested filename differed from the raw filename
    ///     (i.e. at least one sanitization pass changed the name).
    /// </summary>
    [JsonPropertyName("sanitizedFromRaw")]
    [JsonPropertyOrder(4)]
    public bool SanitizedFromRaw { get; init; }

    /// <summary>
    ///     See <see cref="OleListResult.PasswordIgnored" />.
    /// </summary>
    [JsonPropertyName("passwordIgnored")]
    [JsonPropertyOrder(5)]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public PasswordIgnoredNote? PasswordIgnored { get; init; }
}
