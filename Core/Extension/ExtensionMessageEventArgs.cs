namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Event arguments for extension messages.
/// </summary>
public class ExtensionMessageEventArgs : EventArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtensionMessageEventArgs" /> class.
    /// </summary>
    /// <param name="messageType">Type of the message.</param>
    /// <param name="rawJson">Raw JSON content.</param>
    public ExtensionMessageEventArgs(string messageType, string rawJson)
    {
        MessageType = messageType;
        RawJson = rawJson;
    }

    /// <summary>
    ///     Gets the message type.
    /// </summary>
    public string MessageType { get; }

    /// <summary>
    ///     Gets the raw JSON content.
    /// </summary>
    public string RawJson { get; }
}
