using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Acknowledgment message from extension.
/// </summary>
public class ExtensionAckMessage
{
    /// <summary>
    ///     Message type, always "ack".
    /// </summary>
    [JsonPropertyName("type")]
    public string Type { get; set; } = "ack";

    /// <summary>
    ///     Sequence number being acknowledged.
    /// </summary>
    [JsonPropertyName("sequenceNumber")]
    public long SequenceNumber { get; set; }

    /// <summary>
    ///     Status of processing: "processed", "error", etc.
    /// </summary>
    [JsonPropertyName("status")]
    public string Status { get; set; } = "processed";

    /// <summary>
    ///     Optional error message if status is "error".
    /// </summary>
    [JsonPropertyName("error")]
    public string? Error { get; set; }
}
