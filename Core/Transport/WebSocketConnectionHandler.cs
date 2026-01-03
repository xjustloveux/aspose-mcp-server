using System.Diagnostics;
using System.Net.WebSockets;
using System.Text;

namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Handles WebSocket connections by bridging them to a Stdio MCP server process.
///     Each WebSocket connection spawns a dedicated Stdio process for isolation.
///     When MCP SDK adds native WebSocket support, this class can be replaced
///     with the SDK's built-in WebSocket transport by changing Program.cs to use:
///     .WithWebSocketTransport() instead of this custom handler.
/// </summary>
public class WebSocketConnectionHandler
{
    /// <summary>
    ///     Path to the MCP server executable.
    /// </summary>
    private readonly string _executablePath;

    /// <summary>
    ///     Logger instance for this handler.
    /// </summary>
    private readonly ILogger<WebSocketConnectionHandler>? _logger;

    /// <summary>
    ///     Additional command line arguments to pass to the MCP server process.
    /// </summary>
    private readonly string _toolArgs;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WebSocketConnectionHandler" /> class.
    /// </summary>
    /// <param name="executablePath">Path to the MCP server executable.</param>
    /// <param name="toolArgs">Additional command line arguments for the server process.</param>
    /// <param name="loggerFactory">Optional logger factory for creating loggers.</param>
    public WebSocketConnectionHandler(string executablePath, string toolArgs, ILoggerFactory? loggerFactory = null)
    {
        _executablePath = executablePath;
        _toolArgs = toolArgs;
        _logger = loggerFactory?.CreateLogger<WebSocketConnectionHandler>();
    }

    /// <summary>
    ///     Handles a WebSocket connection by bridging to a Stdio process.
    /// </summary>
    /// <param name="webSocket">The WebSocket connection to handle.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    public async Task HandleConnectionAsync(WebSocket webSocket, CancellationToken cancellationToken)
    {
        var connectionId = Guid.NewGuid().ToString("N")[..8];
        _logger?.LogInformation("WebSocket connection {ConnectionId} established", connectionId);

        Process? process = null;

        try
        {
            process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = _executablePath,
                    Arguments = $"{_toolArgs} --stdio",
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardInputEncoding = Encoding.UTF8,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                }
            };

            process.Start();
            _logger?.LogDebug("Started Stdio process {ProcessId} for WebSocket {ConnectionId}", process.Id,
                connectionId);

            var readTask = ReadFromProcessAsync(process, webSocket, connectionId, cancellationToken);

            var writeTask = WriteToProcessAsync(webSocket, process, connectionId, cancellationToken);

            await Task.WhenAny(readTask, writeTask);
        }
        catch (WebSocketException ex) when (ex.WebSocketErrorCode == WebSocketError.ConnectionClosedPrematurely)
        {
            _logger?.LogDebug("WebSocket connection {ConnectionId} closed prematurely", connectionId);
        }
        catch (OperationCanceledException)
        {
            _logger?.LogDebug("WebSocket connection {ConnectionId} cancelled", connectionId);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error handling WebSocket connection {ConnectionId}", connectionId);
        }
        finally
        {
            if (process is { HasExited: false })
                try
                {
                    process.Kill();
                    _logger?.LogDebug("Killed Stdio process for WebSocket {ConnectionId}", connectionId);
                }
                catch
                {
                    // Ignore process kill errors (process may already be terminated)
                }

            process?.Dispose();

            if (webSocket.State == WebSocketState.Open)
                try
                {
                    await webSocket.CloseAsync(
                        WebSocketCloseStatus.NormalClosure,
                        "Connection closed",
                        CancellationToken.None);
                }
                catch
                {
                    // Ignore WebSocket close errors (connection may already be closed)
                }

            _logger?.LogInformation("WebSocket connection {ConnectionId} closed", connectionId);
        }
    }

    /// <summary>
    ///     Reads output from the Stdio process and forwards it to the WebSocket client.
    /// </summary>
    /// <param name="process">The Stdio process to read from.</param>
    /// <param name="webSocket">The WebSocket to send data to.</param>
    /// <param name="connectionId">Connection identifier for logging.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    private async Task ReadFromProcessAsync(Process process, WebSocket webSocket, string connectionId,
        CancellationToken cancellationToken)
    {
        try
        {
            var reader = process.StandardOutput;
            var buffer = new char[4096];

            while (!process.HasExited && webSocket.State == WebSocketState.Open &&
                   !cancellationToken.IsCancellationRequested)
            {
                var count = await reader.ReadAsync(buffer, 0, buffer.Length);
                if (count == 0) break;

                var message = new string(buffer, 0, count);
                var bytes = Encoding.UTF8.GetBytes(message);

                await webSocket.SendAsync(
                    new ArraySegment<byte>(bytes),
                    WebSocketMessageType.Text,
                    true,
                    cancellationToken);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error reading from process for WebSocket {ConnectionId}", connectionId);
        }
    }

    /// <summary>
    ///     Reads messages from the WebSocket client and writes them to the Stdio process.
    /// </summary>
    /// <param name="webSocket">The WebSocket to receive data from.</param>
    /// <param name="process">The Stdio process to write to.</param>
    /// <param name="connectionId">Connection identifier for logging.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    private async Task WriteToProcessAsync(WebSocket webSocket, Process process, string connectionId,
        CancellationToken cancellationToken)
    {
        try
        {
            var buffer = new byte[4096];

            while (!process.HasExited && webSocket.State == WebSocketState.Open &&
                   !cancellationToken.IsCancellationRequested)
            {
                var result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), cancellationToken);

                if (result.MessageType == WebSocketMessageType.Close) break;

                if (result.MessageType == WebSocketMessageType.Text)
                {
                    var message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                    await process.StandardInput.WriteLineAsync(message);
                    await process.StandardInput.FlushAsync(cancellationToken);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error writing to process for WebSocket {ConnectionId}", connectionId);
        }
    }
}