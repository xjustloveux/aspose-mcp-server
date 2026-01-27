namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Transport mode for MCP server
/// </summary>
public enum TransportMode
{
    /// <summary>Standard input/output transport</summary>
    Stdio,

    /// <summary>Streamable HTTP transport (MCP 2025-03-26+)</summary>
    Http,

    /// <summary>WebSocket transport</summary>
    WebSocket
}
