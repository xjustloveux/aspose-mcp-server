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

/// <summary>
///     Transport configuration for MCP server
/// </summary>
public class TransportConfig
{
    /// <summary>
    ///     Transport mode (Stdio, SSE, WebSocket)
    /// </summary>
    public TransportMode Mode { get; set; } = TransportMode.Stdio;

    /// <summary>
    ///     Port number for SSE/WebSocket modes
    /// </summary>
    public int Port { get; set; } = 3000;

    /// <summary>
    ///     Host address for SSE/WebSocket modes
    /// </summary>
    public string Host { get; set; } = "localhost";

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    ///     Command line arguments take precedence over environment variables.
    /// </summary>
    /// <param name="args">Command line arguments</param>
    /// <returns>TransportConfig instance</returns>
    public static TransportConfig LoadFromArgs(string[] args)
    {
        var config = new TransportConfig();
        config.LoadFromEnvironment();
        config.LoadFromCommandLine(args);
        return config;
    }

    /// <summary>
    ///     Loads configuration from environment variables
    /// </summary>
    private void LoadFromEnvironment()
    {
        // Transport mode
        var transport = Environment.GetEnvironmentVariable("ASPOSE_TRANSPORT");
        if (!string.IsNullOrEmpty(transport))
            Mode = transport.ToLower() switch
            {
                "stdio" => TransportMode.Stdio,
                "sse" => TransportMode.Sse,
                "ws" or "websocket" => TransportMode.WebSocket,
                _ => Mode
            };

        // Port
        var portStr = Environment.GetEnvironmentVariable("ASPOSE_PORT");
        if (!string.IsNullOrEmpty(portStr) && int.TryParse(portStr, out var port))
            Port = port;

        // Host
        var host = Environment.GetEnvironmentVariable("ASPOSE_HOST");
        if (!string.IsNullOrEmpty(host))
            Host = host;
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables)
    /// </summary>
    /// <param name="args">Command line arguments</param>
    private void LoadFromCommandLine(string[] args)
    {
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            // Transport mode
            if (arg.Equals("--stdio", StringComparison.OrdinalIgnoreCase))
            {
                Mode = TransportMode.Stdio;
            }
            else if (arg.Equals("--sse", StringComparison.OrdinalIgnoreCase))
            {
                Mode = TransportMode.Sse;
            }
            else if (arg.Equals("--ws", StringComparison.OrdinalIgnoreCase) ||
                     arg.Equals("--websocket", StringComparison.OrdinalIgnoreCase))
            {
                Mode = TransportMode.WebSocket;
            }
            // Port with space separator
            else if (arg.Equals("--port", StringComparison.OrdinalIgnoreCase) &&
                     i + 1 < args.Length && int.TryParse(args[i + 1], out var port1))
            {
                Port = port1;
                i++;
            }
            // Port with : or = separator
            else if (arg.StartsWith("--port:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--port:".Length..], out var port2))
            {
                Port = port2;
            }
            else if (arg.StartsWith("--port=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--port=".Length..], out var port3))
            {
                Port = port3;
            }
            // Host with space separator
            else if (arg.Equals("--host", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                Host = args[i + 1];
                i++;
            }
            // Host with : or = separator
            else if (arg.StartsWith("--host:", StringComparison.OrdinalIgnoreCase))
            {
                Host = arg["--host:".Length..];
            }
            else if (arg.StartsWith("--host=", StringComparison.OrdinalIgnoreCase))
            {
                Host = arg["--host=".Length..];
            }
        }
    }
}