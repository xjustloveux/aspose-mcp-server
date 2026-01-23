namespace AsposeMcpServer.Core.Tracking;

/// <summary>
///     Log output target
/// </summary>
public enum LogTarget
{
    /// <summary>
    ///     Output to stderr (following MCP specification)
    /// </summary>
    Console,

    /// <summary>
    ///     Windows Event Log (Windows only)
    /// </summary>
    EventLog
}
