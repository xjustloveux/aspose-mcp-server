using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Heartbeat pong message from extension.
/// </summary>
public class ExtensionPongMessage
{
    /// <summary>
    ///     Message type, always "pong".
    /// </summary>
    [JsonPropertyName("type")]
    public string Type { get; set; } = "pong";
}
