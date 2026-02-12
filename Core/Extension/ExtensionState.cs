namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Represents the state of an extension instance.
/// </summary>
public enum ExtensionState
{
    /// <summary>
    ///     Extension is not loaded (initial state).
    /// </summary>
    Unloaded,

    /// <summary>
    ///     Extension is starting up.
    /// </summary>
    Starting,

    /// <summary>
    ///     Extension process has started and is performing initialization handshake.
    ///     The process is running but handshake is not yet complete.
    /// </summary>
    Initializing,

    /// <summary>
    ///     Extension is ready and waiting for work.
    /// </summary>
    Idle,

    /// <summary>
    ///     Extension is currently processing a snapshot.
    /// </summary>
    Busy,

    /// <summary>
    ///     Extension encountered an error.
    /// </summary>
    Error,

    /// <summary>
    ///     Extension is shutting down.
    /// </summary>
    Stopping,

    /// <summary>
    ///     Extension process crashed unexpectedly.
    /// </summary>
    Crashed
}
