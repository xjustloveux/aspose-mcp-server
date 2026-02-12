namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Known message types for extension communication.
/// </summary>
/// <remarks>
///     <para>Message flow:</para>
///     <list type="bullet">
///         <item>
///             MCP → Extension: <see cref="Initialize" /> (handshake), <see cref="Initialized" /> (handshake complete),
///             <see cref="Snapshot" />, <see cref="Heartbeat" />, <see cref="SessionClosed" />,
///             <see cref="SessionUnbound" />, <see cref="Shutdown" />, <see cref="Command" />
///         </item>
///         <item>
///             Extension → MCP: <see cref="InitializeResponse" /> (handshake), <see cref="Ack" />,
///             <see cref="Pong" />, <see cref="CommandResult" />
///         </item>
///     </list>
///     <para>
///         Extensions must respond to <see cref="Initialize" /> with <see cref="InitializeResponse" />
///         containing name (required) and version (required). After receiving <see cref="Initialized" />,
///         the extension enters normal operation mode.
///     </para>
///     <para>
///         Extensions should acknowledge snapshots with <see cref="Ack" /> messages
///         containing the same sequence number. Heartbeats should be responded to
///         with <see cref="Pong" /> messages.
///     </para>
/// </remarks>
public static class ExtensionMessageType
{
    /// <summary>
    ///     Server sends to extension to initiate handshake.
    /// </summary>
    public const string Initialize = "initialize";

    /// <summary>
    ///     Extension responds with its metadata during handshake.
    /// </summary>
    public const string InitializeResponse = "initialize_response";

    /// <summary>
    ///     Server confirms handshake completion.
    /// </summary>
    public const string Initialized = "initialized";

    /// <summary>
    ///     Document snapshot message.
    /// </summary>
    public const string Snapshot = "snapshot";

    /// <summary>
    ///     Heartbeat request message.
    /// </summary>
    public const string Heartbeat = "heartbeat";

    /// <summary>
    ///     Session closed notification.
    /// </summary>
    public const string SessionClosed = "session_closed";

    /// <summary>
    ///     Session unbound notification (session still exists but is no longer bound to this extension).
    /// </summary>
    public const string SessionUnbound = "session_unbound";

    /// <summary>
    ///     Shutdown notification.
    /// </summary>
    public const string Shutdown = "shutdown";

    /// <summary>
    ///     Acknowledgment from extension.
    /// </summary>
    public const string Ack = "ack";

    /// <summary>
    ///     Heartbeat response from extension.
    /// </summary>
    public const string Pong = "pong";

    /// <summary>
    ///     Command message to extension.
    /// </summary>
    public const string Command = "command";

    /// <summary>
    ///     Command result from extension.
    /// </summary>
    public const string CommandResult = "command_result";
}
