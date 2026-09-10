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

    /// <summary>Largest number of frames one logical message may be split across.</summary>
    private const int MaxMessageFragments = 10_000;

    /// <summary>
    ///     Largest single line accepted from the child's stdout or stderr, in characters. Reading
    ///     a line has no bound of its own, so a child that emits no newline is an unbounded
    ///     allocation.
    /// </summary>
    private const int MaxChildLineChars = MaxMessageSize;

    /// <summary>Largest total stderr logged for one connection, in characters.</summary>
    private const long MaxStderrChars = 4L * 1024 * 1024;

    /// <summary>
    ///     Number of simultaneous bridged connections allowed when the host does not configure one.
    /// </summary>
    public const int DefaultMaxConcurrentConnections = 32;

    /// <summary>
    ///     Longest a single logical message may take to arrive, counted from its first frame.
    ///     <para>
    ///         The idle deadline is pushed out by every frame that moves, and a frame carrying one
    ///         byte moves. A client sending a byte at a time therefore held a child process, a
    ///         connection slot and the memory of a half-assembled message for as long as it liked
    ///         while never being idle (R3-R02). This deadline is absolute: it does not move.
    ///     </para>
    /// </summary>
    private static readonly TimeSpan MaxMessageAssembly = TimeSpan.FromSeconds(60);

    /// <summary>
    ///     How long the close handshake may take before the socket is aborted.
    ///     <para>
    ///         Waiting on <c>CloseAsync</c> with no cancellation, and doing it before the
    ///         connection slot was released, let an uncooperative peer hold a child process slot
    ///         for as long as it liked (R4-S08).
    ///     </para>
    /// </summary>
    private static readonly TimeSpan CloseHandshakeTimeout = TimeSpan.FromSeconds(5);

    /// <summary>
    ///     Command line arguments passed to each MCP server process, as individual tokens so a
    ///     value containing spaces needs no quoting.
    /// </summary>
    private readonly IReadOnlyList<string> _childArguments;

    /// <summary>
    ///     Bounds how many connections may hold a child process at once. Each connection starts a
    ///     dedicated process, so without a bound a caller can exhaust process handles, memory and
    ///     CPU by opening sockets.
    /// </summary>
    private readonly SemaphoreSlim _connectionSlots;

    /// <summary>
    ///     Path to the MCP server executable.
    /// </summary>
    private readonly string _executablePath;

    /// <summary>
    ///     How long a connection may hold its slot without any WebSocket traffic before it is
    ///     closed and its child process terminated.
    /// </summary>
    private readonly TimeSpan _idleTimeout;

    /// <summary>
    ///     Logger instance for this handler.
    /// </summary>
    private readonly ILogger<WebSocketConnectionHandler>? _logger;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WebSocketConnectionHandler" /> class.
    /// </summary>
    /// <param name="executablePath">Path to the MCP server executable.</param>
    /// <param name="childArguments">Argument tokens for the server process, excluding transport selection.</param>
    /// <param name="loggerFactory">Optional logger factory for creating loggers.</param>
    /// <param name="maxConcurrentConnections">Maximum number of simultaneous bridged connections.</param>
    /// <param name="idleTimeout">
    ///     Idle period after which a connection is closed. Omitting it uses thirty minutes;
    ///     it cannot be disabled, because a connection holds a child process (R4-DOC03).
    /// </param>
    public WebSocketConnectionHandler(string executablePath, IReadOnlyList<string> childArguments,
        ILoggerFactory? loggerFactory = null, int maxConcurrentConnections = DefaultMaxConcurrentConnections,
        TimeSpan? idleTimeout = null)
    {
        _executablePath = executablePath;
        _childArguments = childArguments;
        _logger = loggerFactory?.CreateLogger<WebSocketConnectionHandler>();
        _connectionSlots = new SemaphoreSlim(Math.Max(1, maxConcurrentConnections));
        _idleTimeout = idleTimeout ?? TimeSpan.FromMinutes(30);
    }

    /// <summary>
    ///     Handles a WebSocket connection by bridging to a Stdio process.
    /// </summary>
    /// <param name="webSocket">The WebSocket connection to handle.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <param name="groupId">Optional group ID from authentication (passed to child process).</param>
    /// <param name="userId">Optional user ID from authentication (passed to child process).</param>
    /// <returns>A task that completes when the work is done.</returns>
    public async Task HandleConnectionAsync(
        WebSocket webSocket,
        CancellationToken cancellationToken,
        string? groupId = null,
        string? userId = null)
    {
        var connectionId = Guid.NewGuid().ToString("N")[..8];

        if (!await _connectionSlots.WaitAsync(TimeSpan.Zero, cancellationToken))
        {
            _logger?.LogWarning(
                "Rejected WebSocket connection {ConnectionId}: all {Limit} connection slots are in use",
                connectionId, _connectionSlots.CurrentCount + 1);
            await CloseWithoutBridgingAsync(webSocket, _logger, cancellationToken);
            return;
        }

        _logger?.LogInformation("WebSocket connection {ConnectionId} established", connectionId);

        Process? process = null;
        using var deadline = new RefreshableCancellationDeadline(_idleTimeout, cancellationToken);

        try
        {
            var startInfo = ChildProcessArguments.CreateChildStartInfo(
                _executablePath, _childArguments, groupId, userId);

            process = new Process { StartInfo = startInfo };

            process.Start();
            _logger?.LogDebug("Started Stdio process {ProcessId} for WebSocket {ConnectionId}", process.Id,
                connectionId);

            var readTask = ReadFromProcessAsync(process, webSocket, connectionId, deadline.Refresh,
                deadline.Token);
            var writeTask = WriteToProcessAsync(webSocket, process, connectionId, deadline.Refresh,
                deadline.Token);
            var stderrTask = ReadStderrAsync(process, connectionId, deadline.Token);

            await Task.WhenAny(readTask, writeTask, stderrTask);

            await deadline.CancelAsync();

            await Task.WhenAll(
                readTask.ContinueWith(_ => { }, TaskContinuationOptions.OnlyOnFaulted),
                writeTask.ContinueWith(_ => { }, TaskContinuationOptions.OnlyOnFaulted),
                stderrTask.ContinueWith(_ => { }, TaskContinuationOptions.OnlyOnFaulted)
            ).WaitAsync(TimeSpan.FromSeconds(2), CancellationToken.None);
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
                    process.StandardInput.Close();
                    if (!process.WaitForExit(1000))
                    {
                        // The child is another copy of this server and can have started processes
                        // of its own; killing only the parent leaves those running (R3-R02).
                        process.Kill(true);
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

            // Released before the close handshake, which is the only part of teardown the peer can
            // stall: holding the slot across it let an uncooperative client keep a child process
            // slot out of circulation (R4-S08).
            _connectionSlots.Release();

            await CloseWithTimeoutOrAbortAsync(webSocket, WebSocketCloseStatus.NormalClosure,
                "Connection closed", _logger, connectionId, cancellationToken);

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
    /// <param name="keepAlive">
    ///     Called whenever traffic moves, to push out the connection's idle deadline.
    /// </param>
    /// <returns>A task that completes when the child stops producing output.</returns>
    private async Task ReadFromProcessAsync(Process process, WebSocket webSocket, string connectionId,
        Action keepAlive, CancellationToken cancellationToken)
    {
        try
        {
            var reader = new BoundedLineReader(process.StandardOutput, MaxChildLineChars);

            while (!process.HasExited && webSocket.State == WebSocketState.Open &&
                   !cancellationToken.IsCancellationRequested)
            {
                // One JSON-RPC message per line, read with a bound (R3-R02).
                var line = await reader.ReadLineAsync(cancellationToken);
                if (line == null) break;

                if (reader.LineTooLong)
                {
                    _logger?.LogWarning(
                        "WebSocket {ConnectionId} child wrote a line above {MaxChars} characters; closing the connection",
                        connectionId, MaxChildLineChars);
                    break;
                }

                keepAlive();
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
    /// <returns>A task that completes when the child stops producing diagnostics.</returns>
    private async Task ReadStderrAsync(Process process, string connectionId, CancellationToken cancellationToken)
    {
        try
        {
            var reader = new BoundedLineReader(process.StandardError, MaxChildLineChars);
            var logged = 0L;
            var capReported = false;

            while (!process.HasExited && !cancellationToken.IsCancellationRequested)
            {
                var line = await reader.ReadLineAsync(cancellationToken);
                if (line == null) break;

                // The pipe is still drained past the cap: leaving it unread would block the child
                // rather than protect anything. Only the logging stops (R3-R02).
                if (logged >= MaxStderrChars)
                {
                    if (capReported) continue;
                    _logger?.LogWarning(
                        "WebSocket {ConnectionId} child stderr passed {MaxChars} characters; further output is dropped",
                        connectionId, MaxStderrChars);
                    capReported = true;
                    continue;
                }

                logged += line.Length;
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
    /// <param name="keepAlive">
    ///     Called whenever traffic moves, to push out the connection's idle deadline.
    /// </param>
    /// <returns>
    ///     A task that completes when the socket closes, the child exits, or a message passes its
    ///     assembly deadline (R4-S08).
    /// </returns>
    private Task WriteToProcessAsync(WebSocket webSocket, Process process, string connectionId,
        Action keepAlive, CancellationToken cancellationToken)
    {
        return BridgeToChildStdinAsync(webSocket, process.StandardInput, () => process.HasExited,
            connectionId, keepAlive, cancellationToken);
    }

    /// <summary>
    ///     Reads messages from the WebSocket client and writes them to the child's standard input.
    ///     <para>
    ///         Separate from the <see cref="Process" /> it normally writes to so a fixture can drive
    ///         every branch of the loop — an oversized message, a message that outruns its fragment
    ///         budget, a frame type this endpoint does not carry — with a scripted socket and no
    ///         child process at all. Those branches decide when a connection is closed and when a
    ///         peer gets to push the idle deadline out, and neither was covered (R4-S08).
    ///     </para>
    /// </summary>
    /// <param name="webSocket">The WebSocket to receive data from.</param>
    /// <param name="childStdin">The child's standard input.</param>
    /// <param name="childHasExited">Answers whether the child is still running.</param>
    /// <param name="connectionId">Connection identifier for logging.</param>
    /// <param name="cancellationToken">Cancellation token for the operation.</param>
    /// <param name="keepAlive">
    ///     Called whenever traffic moves, to push out the connection's idle deadline.
    /// </param>
    /// <returns>
    ///     A task that completes when the socket closes, the child exits, or a message passes its
    ///     assembly deadline (R4-S08).
    /// </returns>
    internal async Task BridgeToChildStdinAsync(WebSocket webSocket, TextWriter childStdin,
        Func<bool> childHasExited, string connectionId, Action keepAlive,
        CancellationToken cancellationToken)
    {
        try
        {
            var buffer = new byte[4096];
            using var messageBuffer = new MemoryStream();
            var assembly = new MessageAssemblyBudget(MaxMessageAssembly, MaxMessageFragments);

            while (!childHasExited() && webSocket.State == WebSocketState.Open &&
                   !cancellationToken.IsCancellationRequested)
            {
                // A half-delivered message has a deadline, and waiting for its next frame has to
                // be part of that deadline: checking the budget only when a frame arrived meant a
                // client that sent one non-final frame and stopped was never measured again, while
                // the frame it did send had already pushed the idle deadline out (R4-S08).
                var remaining = assembly.Remaining(DateTimeOffset.UtcNow, MaxMessageAssembly);
                using var receiveCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                if (remaining < MaxMessageAssembly) receiveCts.CancelAfter(remaining);

                WebSocketReceiveResult result;
                try
                {
                    result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), receiveCts.Token);
                }
                catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
                {
                    _logger?.LogWarning(
                        "WebSocket {ConnectionId} stopped part-way through a message; closing the connection",
                        connectionId);
                    break;
                }

                if (result.MessageType == WebSocketMessageType.Close) break;

                // Anything that is not text was dropped on the floor, and the keep-alive above ran
                // before the decision: a peer could hold a connection and its child process open
                // indefinitely with binary frames the server then ignored (R4-S08).
                if (result.MessageType != WebSocketMessageType.Text)
                {
                    _logger?.LogWarning(
                        "WebSocket {ConnectionId} sent a {MessageType} frame; this endpoint carries " +
                        "JSON-RPC text only, closing the connection",
                        connectionId, result.MessageType);
                    await CloseWithTimeoutOrAbortAsync(webSocket, WebSocketCloseStatus.InvalidMessageType,
                        "This endpoint accepts JSON-RPC text frames only", _logger, connectionId,
                        cancellationToken);
                    break;
                }

                // Only a frame this endpoint will act on is evidence the peer is still working.
                keepAlive();

                // Neither a slow message nor a heavily fragmented one is idleness, so neither
                // was caught by the idle deadline.
                if (!assembly.TryAddFragment(DateTimeOffset.UtcNow))
                {
                    _logger?.LogWarning(
                        "WebSocket {ConnectionId} took too long to deliver one message ({Fragments} frames); closing the connection",
                        connectionId, assembly.Fragments);
                    await CloseWithTimeoutOrAbortAsync(webSocket, WebSocketCloseStatus.PolicyViolation,
                        $"A single message must arrive within {MaxMessageAssembly.TotalSeconds} seconds",
                        _logger, connectionId, cancellationToken);
                    break;
                }

                // Check message size limit to prevent DoS
                if (messageBuffer.Length + result.Count > MaxMessageSize)
                {
                    _logger?.LogWarning(
                        "WebSocket {ConnectionId} message exceeds maximum size ({MaxSize} bytes), closing connection",
                        connectionId, MaxMessageSize);
                    await CloseWithTimeoutOrAbortAsync(webSocket, WebSocketCloseStatus.MessageTooBig,
                        $"Message exceeds maximum size of {MaxMessageSize} bytes",
                        _logger, connectionId, cancellationToken);
                    break;
                }

                messageBuffer.Write(buffer, 0, result.Count);

                // Only process when we have the complete message
                if (result.EndOfMessage)
                {
                    var bytes = messageBuffer.ToArray();

                    // The byte cap above bounds what the peer may send; it does not bound what
                    // binding it costs. Measured per thread: ten megabytes of the smallest JSON
                    // elements bind to about 155 MB, a bounded constant; the guard is defence in
                    // depth, and every per-array limit this server has arrives after the objects
                    // exist (R19-RES02). The HTTP route has the same guard as middleware.
                    if (!PayloadShapeGuard.IsWithinBounds(bytes, out var refusal))
                    {
                        _logger?.LogWarning(
                            "WebSocket {ConnectionId} sent a message this server will not bind: {Reason}; closing the connection",
                            connectionId, refusal);
                        await CloseWithTimeoutOrAbortAsync(webSocket, WebSocketCloseStatus.MessageTooBig,
                            refusal, _logger, connectionId, cancellationToken);
                        break;
                    }

                    var message = Encoding.UTF8.GetString(bytes);
                    await childStdin.WriteLineAsync(message);
                    await childStdin.FlushAsync(cancellationToken);
                    messageBuffer.SetLength(0); // Reset buffer for next message
                    assembly.Complete();
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error writing to process for WebSocket {ConnectionId}", connectionId);
        }
    }

    /// <summary>
    ///     Closes a connection that was refused because no slot was free, without starting a child
    ///     process. The caller is told the reason so it can retry rather than treat it as a fault.
    /// </summary>
    /// <param name="webSocket">The accepted socket to close.</param>
    /// <param name="logger">Logger for the outcome, if any.</param>
    /// <param name="cancellationToken">Token cancelled when the request is aborted.</param>
    /// <returns>A task that completes when the work is done.</returns>
    private static Task CloseWithoutBridgingAsync(WebSocket webSocket, ILogger? logger,
        CancellationToken cancellationToken)
    {
        return CloseWithTimeoutOrAbortAsync(webSocket, WebSocketCloseStatus.PolicyViolation,
            "Server is at its connection limit", logger, "refused", cancellationToken);
    }

    /// <summary>
    ///     Closes a socket, giving the peer a bounded time to answer and dropping it if it will not.
    ///     <para>
    ///         Two of the close paths passed the request's own token instead, which a peer that
    ///         never answers the handshake does not cancel: the close then waited as long as the
    ///         request lived, holding the connection open exactly where the server had decided to
    ///         refuse it (R4-S08). Every path uses this one, so the deadline cannot be forgotten
    ///         at a new call site.
    ///     </para>
    /// </summary>
    /// <param name="webSocket">The socket to close.</param>
    /// <param name="status">The close status to report to the peer.</param>
    /// <param name="description">The reason to report to the peer.</param>
    /// <param name="logger">Logger for the outcome, if any.</param>
    /// <param name="connectionId">Connection identifier for logging.</param>
    /// <param name="cancellationToken">Token cancelled when the request is aborted.</param>
    /// <returns>A task that completes when the socket is closed or aborted.</returns>
    private static async Task CloseWithTimeoutOrAbortAsync(WebSocket webSocket,
        WebSocketCloseStatus status, string description, ILogger? logger, string connectionId,
        CancellationToken cancellationToken)
    {
        if (webSocket.State != WebSocketState.Open) return;

        try
        {
            using var closeCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            closeCts.CancelAfter(CloseHandshakeTimeout);
            await webSocket.CloseAsync(status, description, closeCts.Token);
        }
        catch (Exception ex)
        {
            // A peer that will not complete the handshake is dropped rather than waited for; the
            // socket is going away either way.
            logger?.LogDebug(ex, "WebSocket {ConnectionId} close handshake did not complete",
                connectionId);
            webSocket.Abort();
        }
    }
}

/// <summary>
///     Owns a linked cancellation deadline whose refresh callback may finish after its connection
///     handler has begun disposing resources.
/// </summary>
internal sealed class RefreshableCancellationDeadline : IDisposable
{
    private readonly object _gate = new();
    private readonly CancellationTokenSource _source;
    private readonly TimeSpan _timeout;
    private bool _disposed;

    /// <summary>Creates and starts a deadline linked to the request cancellation token.</summary>
    /// <param name="cancellationToken">The request cancellation token.</param>
    /// <param name="timeout">How far each refresh moves the idle deadline.</param>
    public RefreshableCancellationDeadline(TimeSpan timeout, CancellationToken cancellationToken)
    {
        _timeout = timeout;
        _source = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        Token = _source.Token;
        Refresh();
    }

    /// <summary>The token cancelled by the request or when the idle deadline expires.</summary>
    public CancellationToken Token { get; }

    /// <inheritdoc />
    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            _source.Dispose();
        }
    }

    /// <summary>Moves the idle deadline after traffic, or does nothing once disposal has won the race.</summary>
    public void Refresh()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _source.CancelAfter(_timeout);
        }
    }

    /// <summary>Cancels the linked operation without blocking synchronous cancellation callbacks.</summary>
    /// <returns>A task that completes after cancellation callbacks finish.</returns>
    public Task CancelAsync()
    {
        lock (_gate)
        {
            return _disposed ? Task.CompletedTask : _source.CancelAsync();
        }
    }
}
