using System.Diagnostics;
using System.Net.WebSockets;
using System.Reflection;
using System.Text;
using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Core.Transport;

/// <summary>
///     What the bridge does when a connection has to end, and what counts as the peer being alive.
///     <para>
///         Two close paths passed the request's own cancellation token, which a peer that never
///         answers the handshake never cancels, so the refusal waited as long as the request lived.
///         And the idle deadline was pushed out before the frame type was looked at, so a peer could
///         hold a connection and its child process open indefinitely with binary frames the server
///         went on to ignore (R4-S08).
///     </para>
/// </summary>
public class WebSocketCloseAndKeepAliveTests
{
    private static readonly TimeSpan CloseHandshakeTimeout =
        (TimeSpan)typeof(WebSocketConnectionHandler)
            .GetField("CloseHandshakeTimeout", BindingFlags.NonPublic | BindingFlags.Static)!
            .GetValue(null)!;

    private static readonly int MaxMessageFragments =
        (int)typeof(WebSocketConnectionHandler)
            .GetField("MaxMessageFragments", BindingFlags.NonPublic | BindingFlags.Static)!
            .GetValue(null)!;

    private static readonly int MaxMessageSize =
        (int)typeof(WebSocketConnectionHandler)
            .GetField("MaxMessageSize", BindingFlags.NonPublic | BindingFlags.Static)!
            .GetValue(null)!;

    [Fact]
    public void RefreshingAnIdleDeadlineAfterDisposal_ShouldBeHarmless()
    {
        var deadline = new RefreshableCancellationDeadline(
            TimeSpan.FromMinutes(1), CancellationToken.None);
        deadline.Dispose();

        var error = Record.Exception(deadline.Refresh);

        Assert.Null(error);
    }

    /// <summary>
    ///     Waits for work that is supposed to have its own deadline, and fails if it has not.
    ///     <para>
    ///         Written this way on purpose: the defect these two fixtures cover is a close that
    ///         waits on a token an unresponsive peer never cancels, so an unbounded await here
    ///         would hang the run instead of reporting it.
    ///     </para>
    /// </summary>
    /// <param name="start">Starts the work. Called after the clock, so the wait is measured whole.</param>
    /// <param name="what">What is being waited for, for the failure message.</param>
    /// <returns>How long the work took.</returns>
    private static async Task<TimeSpan> WithinTheHandshakeDeadline(Func<Task> start, string what)
    {
        // The work is started here rather than passed in already running: taking the timestamp
        // after the task had begun left the setup out of the measurement, and under a loaded
        // machine that read back as the close giving up a few milliseconds early.
        var began = Stopwatch.StartNew();
        var work = start();
        var bound = CloseHandshakeTimeout * 3;

        Assert.True(await Task.WhenAny(work, Task.Delay(bound)) == work,
            $"{what} was still waiting after {bound.TotalSeconds:0}s, past its own " +
            $"{CloseHandshakeTimeout.TotalSeconds:0}s handshake deadline");

        await work;
        return began.Elapsed;
    }

    /// <summary>
    ///     Runs the bridge against a scripted socket.
    /// </summary>
    /// <param name="socket">The socket to drive.</param>
    /// <param name="keepAlives">Receives one entry per idle-deadline refresh.</param>
    /// <returns>What the peer's frames delivered to the child's standard input.</returns>
    private static async Task<string> Bridge(ScriptedWebSocket socket, List<int> keepAlives)
    {
        var handler = new WebSocketConnectionHandler("/no/such/executable", ["--stdio"]);
        var stdin = new StringWriter();

        await handler.BridgeToChildStdinAsync(socket, stdin, () => false, "fixture",
            () => keepAlives.Add(keepAlives.Count), CancellationToken.None);

        return stdin.ToString();
    }

    [Fact]
    public async Task ATextMessage_ShouldReachTheChildAndRefreshTheDeadline()
    {
        var socket = new ScriptedWebSocket();
        socket.Text("{\"jsonrpc\":\"2.0\"}");
        socket.Close();

        var keepAlives = new List<int>();
        var forwarded = await Bridge(socket, keepAlives);

        Assert.Equal("{\"jsonrpc\":\"2.0\"}" + Environment.NewLine, forwarded);
        Assert.Single(keepAlives);
        Assert.Empty(socket.Closes);
    }

    /// <summary>
    ///     A binary frame is not something this endpoint carries, so it is neither forwarded nor
    ///     treated as the peer doing work. It used to be both ignored and credited with keeping the
    ///     connection alive, which is a connection that never times out and never does anything.
    /// </summary>
    [Fact]
    public async Task ABinaryFrame_ShouldCloseTheConnectionWithoutRefreshingTheDeadline()
    {
        var socket = new ScriptedWebSocket();
        socket.Binary([1, 2, 3]);
        socket.Text("this must never be read");

        var keepAlives = new List<int>();
        var forwarded = await Bridge(socket, keepAlives);

        Assert.Equal("", forwarded);
        Assert.Empty(keepAlives);
        Assert.Equal(WebSocketCloseStatus.InvalidMessageType, Assert.Single(socket.Closes).Status);
    }

    [Fact]
    public async Task AMessageOverTheSizeLimit_ShouldBeRefusedRatherThanTruncated()
    {
        var socket = new ScriptedWebSocket();
        var chunk = new string('x', 4096);
        for (var written = 0; written <= MaxMessageSize; written += chunk.Length)
            socket.TextFragment(chunk);

        var keepAlives = new List<int>();
        var forwarded = await Bridge(socket, keepAlives);

        Assert.Equal("", forwarded);
        Assert.Equal(WebSocketCloseStatus.MessageTooBig, Assert.Single(socket.Closes).Status);
    }

    [Fact]
    public async Task AMessageOverTheFragmentBudget_ShouldBeRefused()
    {
        var socket = new ScriptedWebSocket();
        for (var i = 0; i <= MaxMessageFragments; i++)
            socket.TextFragment("x");

        var keepAlives = new List<int>();
        var forwarded = await Bridge(socket, keepAlives);

        Assert.Equal("", forwarded);
        Assert.Equal(WebSocketCloseStatus.PolicyViolation, Assert.Single(socket.Closes).Status);
    }

    /// <summary>
    ///     A peer that accepts the close frame and then says nothing must not hold the server. The
    ///     handshake has its own deadline, and the socket is dropped when it passes.
    /// </summary>
    [Fact]
    public async Task APeerThatNeverCompletesTheHandshake_ShouldBeDroppedAtTheDeadline()
    {
        var socket = new ScriptedWebSocket { StallTheCloseHandshake = true };
        socket.Binary([9]);

        var waited = await WithinTheHandshakeDeadline(() => Bridge(socket, []), "the close");

        Assert.True(socket.Aborted, "the socket was left to a peer that never answered");

        // A CancelAfter timer is coarse, so the deadline is checked with the slack the platform
        // itself has rather than to the millisecond.
        Assert.True(waited >= CloseHandshakeTimeout - TimeSpan.FromMilliseconds(250),
            $"the close gave up after {waited.TotalSeconds:0.00}s, before the "
            + $"{CloseHandshakeTimeout.TotalSeconds:0}s deadline it promises the peer");
    }

    /// <summary>
    ///     The refusal path — no slot free — closes without ever starting a child process, and it is
    ///     one of the two that passed the request token instead of the handshake deadline.
    /// </summary>
    [Fact]
    public async Task RefusingAConnection_ShouldNotWaitOnTheRequestToken()
    {
        var socket = new ScriptedWebSocket { StallTheCloseHandshake = true };

        await WithinTheHandshakeDeadline(() => (Task)typeof(WebSocketConnectionHandler)
            .GetMethod("CloseWithoutBridgingAsync", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [socket, null, CancellationToken.None])!, "the refusal");

        Assert.Equal(WebSocketCloseStatus.PolicyViolation, Assert.Single(socket.Closes).Status);
        Assert.True(socket.Aborted, "a peer that will not answer must be dropped, not waited for");
    }

    /// <summary>
    ///     A socket whose frames are decided in advance, so every branch of the receive loop is
    ///     reachable without a peer or a child process.
    /// </summary>
    private sealed class ScriptedWebSocket : WebSocket
    {
        private readonly Queue<(WebSocketMessageType Type, byte[] Payload, bool EndOfMessage)> _frames = new();

        /// <summary>Close frames this socket was asked to send, in order.</summary>
        public List<(WebSocketCloseStatus? Status, string? Description)> Closes { get; } = [];

        /// <summary>Whether the socket was dropped rather than closed politely.</summary>
        public bool Aborted { get; private set; }

        /// <summary>Makes <see cref="CloseAsync" /> wait for its caller's token instead of answering.</summary>
        public bool StallTheCloseHandshake { get; init; }

        /// <inheritdoc />
        public override WebSocketCloseStatus? CloseStatus => null;

        /// <inheritdoc />
        public override string? CloseStatusDescription => null;

        /// <inheritdoc />
        public override WebSocketState State { get; } = WebSocketState.Open;

        /// <inheritdoc />
        public override string? SubProtocol => null;

        /// <summary>Queues one complete text message.</summary>
        /// <param name="payload">Its content.</param>
        public void Text(string payload)
        {
            _frames.Enqueue((WebSocketMessageType.Text, Encoding.UTF8.GetBytes(payload), true));
        }

        /// <summary>Queues one non-final text frame.</summary>
        /// <param name="payload">Its content.</param>
        public void TextFragment(string payload)
        {
            _frames.Enqueue((WebSocketMessageType.Text, Encoding.UTF8.GetBytes(payload), false));
        }

        /// <summary>Queues one binary frame, which this endpoint does not carry.</summary>
        /// <param name="payload">Its content.</param>
        public void Binary(byte[] payload)
        {
            _frames.Enqueue((WebSocketMessageType.Binary, payload, true));
        }

        /// <summary>Queues the peer's own close frame.</summary>
        public void Close()
        {
            _frames.Enqueue((WebSocketMessageType.Close, [], true));
        }

        /// <inheritdoc />
        public override void Abort()
        {
            Aborted = true;
        }

        /// <inheritdoc />
        public override async Task CloseAsync(WebSocketCloseStatus closeStatus, string? statusDescription,
            CancellationToken cancellationToken)
        {
            Closes.Add((closeStatus, statusDescription));
            if (StallTheCloseHandshake) await Task.Delay(Timeout.Infinite, cancellationToken);
        }

        /// <inheritdoc />
        public override Task CloseOutputAsync(WebSocketCloseStatus closeStatus, string? statusDescription,
            CancellationToken cancellationToken)
        {
            return CloseAsync(closeStatus, statusDescription, cancellationToken);
        }

        /// <inheritdoc />
        public override void Dispose()
        {
        }

        /// <inheritdoc />
        public override Task<WebSocketReceiveResult> ReceiveAsync(ArraySegment<byte> buffer,
            CancellationToken cancellationToken)
        {
            if (_frames.Count == 0)
                return Task.FromResult(new WebSocketReceiveResult(0, WebSocketMessageType.Close, true));

            var (type, payload, endOfMessage) = _frames.Dequeue();
            payload.CopyTo(buffer.Array!, buffer.Offset);
            return Task.FromResult(new WebSocketReceiveResult(payload.Length, type, endOfMessage));
        }

        /// <inheritdoc />
        public override Task SendAsync(ArraySegment<byte> buffer, WebSocketMessageType messageType,
            bool endOfMessage, CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }
    }
}
