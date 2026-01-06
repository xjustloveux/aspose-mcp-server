using System.Net;

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
        config.Validate();
        return config;
    }

    /// <summary>
    ///     Validates the configuration values
    /// </summary>
    private void Validate()
    {
        if (Port is < 1 or > 65535)
        {
            Console.Error.WriteLine($"[WARN] Invalid port {Port}, using default 3000");
            Port = 3000;
        }

        if (!IsValidHost(Host))
        {
            Console.Error.WriteLine($"[WARN] Invalid host '{Host}', using default 'localhost'");
            Host = "localhost";
        }
    }

    /// <summary>
    ///     Checks if the host value is valid
    /// </summary>
    /// <param name="host">The host value to validate</param>
    /// <returns>True if valid, false otherwise</returns>
    private static bool IsValidHost(string host)
    {
        if (string.IsNullOrWhiteSpace(host))
            return false;

        if (host is "localhost" or "0.0.0.0" or "*")
            return true;

        return IPAddress.TryParse(host, out _);
    }

    /// <summary>
    ///     Loads configuration from environment variables
    /// </summary>
    private void LoadFromEnvironment()
    {
        var transport = Environment.GetEnvironmentVariable("ASPOSE_TRANSPORT");
        if (!string.IsNullOrEmpty(transport))
            Mode = transport.ToLower() switch
            {
                "stdio" => TransportMode.Stdio,
                "sse" => TransportMode.Sse,
                "ws" or "websocket" => TransportMode.WebSocket,
                _ => Mode
            };

        var portStr = Environment.GetEnvironmentVariable("ASPOSE_PORT");
        if (!string.IsNullOrEmpty(portStr) && int.TryParse(portStr, out var port))
            Port = port;

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
            else if (arg.Equals("--port", StringComparison.OrdinalIgnoreCase) &&
                     i + 1 < args.Length && int.TryParse(args[i + 1], out var port1))
            {
                Port = port1;
                i++;
            }
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
            else if (arg.Equals("--host", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                Host = args[i + 1];
                i++;
            }
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