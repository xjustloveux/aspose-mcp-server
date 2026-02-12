using System.IO.Hashing;
using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Metadata sent to extensions with each message.
/// </summary>
public class ExtensionMetadata
{
    /// <summary>
    ///     Protocol version for compatibility checking.
    /// </summary>
    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = "1.0";

    /// <summary>
    ///     Type of message: "snapshot", "heartbeat", "session_closed", "shutdown", or "command".
    /// </summary>
    [JsonPropertyName("type")]
    public string Type { get; set; } = "snapshot";

    /// <summary>
    ///     Session ID for the document.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public string SessionId { get; set; } = string.Empty;

    /// <summary>
    ///     Type of document: "word", "excel", "powerpoint", or "pdf".
    /// </summary>
    [JsonPropertyName("documentType")]
    public string DocumentType { get; set; } = string.Empty;

    /// <summary>
    ///     Original file path of the document (if available).
    /// </summary>
    [JsonPropertyName("originalPath")]
    public string? OriginalPath { get; set; }

    /// <summary>
    ///     Output format of the converted document (e.g., "pdf", "html", "png").
    /// </summary>
    [JsonPropertyName("outputFormat")]
    public string OutputFormat { get; set; } = string.Empty;

    /// <summary>
    ///     MIME type of the output format.
    /// </summary>
    [JsonPropertyName("mimeType")]
    public string MimeType { get; set; } = string.Empty;

    /// <summary>
    ///     Timestamp when the snapshot was created.
    /// </summary>
    [JsonPropertyName("timestamp")]
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;

    /// <summary>
    ///     Sequence number for ordering and ack tracking.
    /// </summary>
    [JsonPropertyName("sequenceNumber")]
    public long SequenceNumber { get; set; }

    /// <summary>
    ///     Transport mode used: "mmap", "stdin", or "file".
    /// </summary>
    [JsonPropertyName("transportMode")]
    public string TransportMode { get; set; } = "file";

    /// <summary>
    ///     Size of the data in bytes.
    /// </summary>
    [JsonPropertyName("dataSize")]
    public long DataSize { get; set; }

    /// <summary>
    ///     CRC32 checksum of the data for integrity verification.
    ///     Extensions should verify this to detect partial/corrupted data.
    /// </summary>
    [JsonPropertyName("checksum")]
    public uint Checksum { get; set; }

    /// <summary>
    ///     Name of the memory-mapped file (for mmap transport mode).
    /// </summary>
    [JsonPropertyName("mmapName")]
    public string? MmapName { get; set; }

    /// <summary>
    ///     Path to the temporary file (for file transport mode).
    /// </summary>
    [JsonPropertyName("filePath")]
    public string? FilePath { get; set; }

    /// <summary>
    ///     Session owner information for isolation and classification.
    ///     Null when IsolationMode=None.
    /// </summary>
    [JsonPropertyName("owner")]
    public SessionOwner? Owner { get; set; }

    /// <summary>
    ///     Custom data for extension-specific use.
    /// </summary>
    [JsonPropertyName("customData")]
    public Dictionary<string, object>? CustomData { get; set; }

    /// <summary>
    ///     Unique identifier for command messages, used to correlate requests and responses.
    ///     Only set when MessageType is "command".
    /// </summary>
    [JsonPropertyName("commandId")]
    public string? CommandId { get; set; }

    /// <summary>
    ///     Type of command being sent (e.g., "highlight", "navigate", "export").
    ///     Only set when MessageType is "command".
    /// </summary>
    [JsonPropertyName("commandType")]
    public string? CommandType { get; set; }

    /// <summary>
    ///     Command payload containing command-specific parameters.
    ///     Only set when MessageType is "command".
    /// </summary>
    [JsonPropertyName("commandPayload")]
    public Dictionary<string, object>? CommandPayload { get; set; }

    /// <summary>
    ///     Verifies the data integrity by comparing the checksum.
    /// </summary>
    /// <param name="data">The data to verify.</param>
    /// <returns>True if the checksum matches, false otherwise.</returns>
    /// <remarks>
    ///     <para>Edge cases:</para>
    ///     <list type="bullet">
    ///         <item>Null data: returns true only if Checksum is 0</item>
    ///         <item>Empty data (length 0): returns true only if Checksum is 0</item>
    ///         <item>Non-empty data: computes CRC32 and compares</item>
    ///     </list>
    ///     <para>
    ///         For complete validation, use <see cref="VerifyData" /> which also checks
    ///         that the data length matches <see cref="DataSize" />.
    ///     </para>
    /// </remarks>
    public bool VerifyChecksum(byte[]? data)
    {
        if (data is null or { Length: 0 })
            return Checksum == 0;

        var computedChecksum = Crc32.HashToUInt32(data);
        return computedChecksum == Checksum;
    }

    /// <summary>
    ///     Verifies both data size and checksum.
    /// </summary>
    /// <param name="data">The data to verify.</param>
    /// <returns>A result indicating success or the type of failure.</returns>
    public DataVerificationResult VerifyData(byte[]? data)
    {
        if (data is null)
            return DataVerificationResult.NullData;

        if (data.Length != DataSize)
            return DataVerificationResult.SizeMismatch;

        if (!VerifyChecksum(data))
            return DataVerificationResult.ChecksumMismatch;

        return DataVerificationResult.Valid;
    }

    /// <summary>
    ///     Creates a command metadata instance.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="commandType">Type of command.</param>
    /// <param name="payload">Command payload.</param>
    /// <returns>A new ExtensionMetadata configured for command message.</returns>
    public static ExtensionMetadata CreateCommand(
        string sessionId,
        string commandType,
        Dictionary<string, object>? payload = null)
    {
        return new ExtensionMetadata
        {
            Type = "command",
            SessionId = sessionId,
            CommandId = Guid.NewGuid().ToString("N"),
            CommandType = commandType,
            CommandPayload = payload,
            Timestamp = DateTime.UtcNow
        };
    }
}
