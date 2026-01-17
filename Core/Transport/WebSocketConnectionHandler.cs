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
    ///     Maximum allowed message size (10 MB) to prevent DoS attacks.
    /// </summary>
    private const int MaxMessageSize = 10 * 1024 * 1024;

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
    /// <param name="groupId">Optional group ID from authentication (passed to child process).</param>
    /// <param name="userId">Optional user ID from authentication (passed to child process).</param>
    public async Task HandleConnectionAsync(
        WebSocket webSocket,
        CancellationToken cancellationToken,
        string? groupId = null,
        string? userId = null)
    {
        var connectionId = Guid.NewGuid().ToString("N")[..8];
        _logger?.LogInformation("WebSocket connection {ConnectionId} established", connectionId);

        Process? process = null;
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        try
        {
            var startInfo = new ProcessStartInfo
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
            };

            if (!string.IsNullOrEmpty(groupId))
                startInfo.Environment["ASPOSE_SESSION_GROUP_ID"] = groupId;
            if (!string.IsNullOrEmpty(userId))
                startInfo.Environment["ASPOSE_SESSION_USER_ID"] = userId;

            process = new Process { StartInfo = startInfo };

            process.Start();
            _logger?.LogDebug("Started Stdio process {ProcessId} for WebSocket {ConnectionId}", process.Id,
                connectionId);

            var readTask = ReadFromProcessAsync(process, webSocket, connectionId, linkedCts.Token);
            var writeTask = WriteToProcessAsync(webSocket, process, connectionId, linkedCts.Token);
            var stderrTask = ReadStderrAsync(process, connectionId, linkedCts.Token);

            await Task.WhenAny(readTask, writeTask, stderrTask);

            await linkedCts.CancelAsync();

            await Task.WhenAll(
                readTask.ContinueWith(_ => { }, TaskContinuationOptions.OnlyOnFaulted),
                writeTask.ContinueWith(_ => { }, TaskContinuationOptions.OnlyOnFaulted),
                stderrTask.ContinueWith(_ => { }, TaskContinuationOptions.OnlyOnFaulted)
                // ReSharper disable once MethodSupportsCancellation - Timeout-only wait during cleanup, no cancellation needed
            ).WaitAsync(TimeSpan.FromSeconds(2));
        }
        catch (WebSocketException ex) when (ex.WebSocketErrorCode == WebSocketError.ConnectionClosedPrematurely)
        {
            _logger?.LogDebug( // NOSONAR S6667 - Structured logging with placeholders is correct pattern
                "WebSocket connection {ConnectionId} closed prematurely",
                connectionId);
        }
        catch (OperationCanceledException)
        {
            _logger?.LogDebug( // NOSONAR S6667 - Structured logging with placeholders is correct pattern
                "WebSocket connection {ConnectionId} cancelled",
                connectionId);
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
                    process.StandardInput.Close();
                    if (!process.WaitForExit(1000))
                    {
                        process.Kill();
                        _logger?.LogDebug("Killed Stdio process for WebSocket {ConnectionId}", connectionId);
                    }
                    else
                    {
                        _logger?.LogDebug("Stdio process exited gracefully for WebSocket {ConnectionId}", connectionId);
                    }
                }
                catch
                {
                    // Ignore process termination errors
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
                    // Ignore WebSocket close errors
                }

            _logger?.LogInformation("WebSocket connection {ConnectionId} closed", connectionId);
        }
    }

    /// <summary>
    ///     Reads output from the Stdio process and forwards it to the WebSocket client.
    ///     Uses ReadLineAsync to ensure complete JSON-RPC messages are sent.
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

            while (!process.HasExited && webSocket.State == WebSocketState.Open &&
                   !cancellationToken.IsCancellationRequested)
            {
                // Use ReadLineAsync to read complete JSON-RPC messages (one per line)
                var line = await reader.ReadLineAsync(cancellationToken);
                if (line == null) break;

                var bytes = Encoding.UTF8.GetBytes(line);

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
    ///     Reads stderr from the Stdio process and logs it.
    ///     This prevents the stderr buffer from filling up and blocking the process.
    /// </summary>
    /// <param name="process">The Stdio process to read stderr from.</param>
    /// <param name="connectionId">Connection identifier for logging.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    private async Task ReadStderrAsync(Process process, string connectionId, CancellationToken cancellationToken)
    {
        try
        {
            var reader = process.StandardError;

            while (!process.HasExited && !cancellationToken.IsCancellationRequested)
            {
                var line = await reader.ReadLineAsync(cancellationToken);
                if (line == null) break;

                _logger?.LogDebug("WebSocket {ConnectionId} stderr: {Line}", connectionId, line);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error reading stderr for WebSocket {ConnectionId}", connectionId);
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
            using var messageBuffer = new MemoryStream();

            while (!process.HasExited && webSocket.State == WebSocketState.Open &&
                   !cancellationToken.IsCancellationRequested)
            {
                var result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), cancellationToken);

                if (result.MessageType == WebSocketMessageType.Close) break;

                if (result.MessageType == WebSocketMessageType.Text)
                {
                    // Check message size limit to prevent DoS
                    if (messageBuffer.Length + result.Count > MaxMessageSize)
                    {
                        _logger?.LogWarning(
                            "WebSocket {ConnectionId} message exceeds maximum size ({MaxSize} bytes), closing connection",
                            connectionId, MaxMessageSize);
                        await webSocket.CloseAsync(
                            WebSocketCloseStatus.MessageTooBig,
                            $"Message exceeds maximum size of {MaxMessageSize} bytes",
                            cancellationToken);
                        break;
                    }

                    // Accumulate message fragments - NOSONAR S6966 - MemoryStream.Write is effectively synchronous
                    messageBuffer.Write(buffer, 0, result.Count);

                    // Only process when we have the complete message
                    if (result.EndOfMessage)
                    {
                        var message = Encoding.UTF8.GetString(messageBuffer.ToArray());
                        await process.StandardInput.WriteLineAsync(message);
                        await process.StandardInput.FlushAsync(cancellationToken);
                        messageBuffer.SetLength(0); // Reset buffer for next message
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error writing to process for WebSocket {ConnectionId}", connectionId);
        }
    }
}
