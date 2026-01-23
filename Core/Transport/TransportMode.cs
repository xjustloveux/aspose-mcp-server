namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Transport mode for MCP server
/// </summary>
public enum TransportMode
{
    /// <summary>Standard input/output transport</summary>
    Stdio,

    /// <summary>Server-Sent Events transport</summary>
    Sse,

    /// <summary>WebSocket transport</summary>
    WebSocket
}
