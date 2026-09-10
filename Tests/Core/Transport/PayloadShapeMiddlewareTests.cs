using System.Text;
using AsposeMcpServer.Core.Transport;
using Microsoft.AspNetCore.Http;

namespace AsposeMcpServer.Tests.Core.Transport;

/// <summary>
///     R20-RES02: the HTTP route gets the same byte and shape bound as WebSocket.
///     <para>
///         Driven through <see cref="DefaultHttpContext" /> rather than a hosted server: the
///         property is what the middleware does with a body, and this is the smallest thing that
///         has one. It is also the honest scope of the evidence — the middleware is registered in
///         <c>HostFactory.ConfigureMiddleware</c> ahead of <c>MapMcp</c>, and that wiring is read
///         from the source, not exercised by a real listener here.
///     </para>
/// </summary>
public class PayloadShapeMiddlewareTests
{
    /// <summary>Runs the middleware over one body and reports what happened.</summary>
    /// <param name="body">The request body.</param>
    /// <returns>The status code and whether the rest of the pipeline was reached.</returns>
    private static async Task<(int Status, bool Reached)> Run(byte[] body)
    {
        var reached = false;
        var middleware = new PayloadShapeMiddleware(_ =>
        {
            reached = true;
            return Task.CompletedTask;
        });

        var context = new DefaultHttpContext
        {
            Request =
            {
                Method = "POST",
                Path = "/mcp",
                ContentType = "application/json",
                ContentLength = body.Length,
                Body = new MemoryStream(body)
            },
            Response = { Body = new MemoryStream() }
        };

        await middleware.InvokeAsync(context);
        return (context.Response.StatusCode, reached);
    }

    [Fact]
    public async Task AnOrdinaryRequest_ShouldReachTheEndpoint()
    {
        var (status, reached) = await Run(
            "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\"}"u8.ToArray());

        Assert.True(reached);
        Assert.Equal(StatusCodes.Status200OK, status);
    }

    [Fact]
    public async Task ABodyOverTheByteCap_ShouldBeRefusedBeforeItIsBound()
    {
        var body = new byte[PayloadShapeMiddleware.MaxBodyBytes + 1];
        Array.Fill(body, (byte)' ');

        var (status, reached) = await Run(body);

        Assert.False(reached);
        Assert.Equal(StatusCodes.Status413PayloadTooLarge, status);
    }

    [Fact]
    public async Task AShapeTheServerWillNotBind_ShouldBeRefusedBeforeItIsBound()
    {
        // The same payload the amplification benchmark measures: many tiny values.
        var builder = new StringBuilder("[");
        for (var i = 0; i < PayloadShapeGuard.MaxValues + 1; i++) builder.Append("0,");
        builder.Append("0]");

        var (status, reached) = await Run(Encoding.UTF8.GetBytes(builder.ToString()));

        Assert.False(reached);
        Assert.Equal(StatusCodes.Status400BadRequest, status);
    }

    [Fact]
    public async Task TheBody_ShouldStillBeReadableByTheEndpointAfterTheCheck()
    {
        // Buffered and rewound: a guard that consumed the body would pass an empty request on.
        string? seen = null;
        var middleware = new PayloadShapeMiddleware(async ctx =>
        {
            using var reader = new StreamReader(ctx.Request.Body);
            seen = await reader.ReadToEndAsync();
        });

        var context = new DefaultHttpContext();
        var body = "{\"jsonrpc\":\"2.0\",\"id\":7,\"method\":\"ping\"}"u8.ToArray();
        context.Request.Method = "POST";
        context.Request.ContentLength = body.Length;
        context.Request.Body = new MemoryStream(body);
        context.Response.Body = new MemoryStream();

        await middleware.InvokeAsync(context);

        Assert.Equal(Encoding.UTF8.GetString(body), seen);
    }

    [Fact]
    public async Task ABodyOfUnknownLengthOverTheCap_ShouldBeRefusedWithoutBeingHeldWhole()
    {
        // R21-RES04. With no Content-Length the earlier version copied the whole body and only
        // then measured it; on a host without the size feature that copy was unbounded.
        // A body three times the cap that reports no length, behind a stream that counts what
        // was read from it. Status alone cannot tell the two versions apart: copying the whole
        // body and then measuring it also ends in 413. What the old version cannot do is stop
        // early.
        var counting = new CountingStream(PayloadShapeMiddleware.MaxBodyBytes * 3);

        var reached = false;
        var middleware = new PayloadShapeMiddleware(_ =>
        {
            reached = true;
            return Task.CompletedTask;
        });

        var context = new DefaultHttpContext
        {
            Request = { Method = "POST", ContentLength = null, Body = counting },
            Response = { Body = new MemoryStream() }
        };

        await middleware.InvokeAsync(context);

        Assert.False(reached);
        Assert.Equal(StatusCodes.Status413PayloadTooLarge, context.Response.StatusCode);
        Assert.True(counting.BytesRead <= PayloadShapeMiddleware.MaxBodyBytes + 2 * 81920,
            $"the middleware read {counting.BytesRead:N0} bytes of a body it was meant to stop reading at {PayloadShapeMiddleware.MaxBodyBytes:N0}");
    }

    [Fact]
    public async Task AShapeRefusal_ShouldWriteWithTheRequestCancellationToken()
    {
        var middleware = new PayloadShapeMiddleware(_ => Task.CompletedTask);
        using var requestLifetime = new CancellationTokenSource();
        var json = new string('[', PayloadShapeGuard.MaxDepth + 1) + "0"
                                                                   + new string(']', PayloadShapeGuard.MaxDepth + 1);
        var body = Encoding.UTF8.GetBytes(json);
        var context = new DefaultHttpContext
        {
            Request =
            {
                ContentLength = body.Length,
                Body = new CancelOnEndStream(body, requestLifetime)
            },
            Response = { Body = new MemoryStream() },
            RequestAborted = requestLifetime.Token
        };

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => middleware.InvokeAsync(context));
    }

    [Fact]
    public async Task ASizeRefusal_ShouldWriteWithTheRequestCancellationToken()
    {
        var middleware = new PayloadShapeMiddleware(_ => Task.CompletedTask);
        using var requestLifetime = new CancellationTokenSource();
        var context = new DefaultHttpContext
        {
            Request =
            {
                ContentLength = PayloadShapeMiddleware.MaxBodyBytes + 1,
                Body = Stream.Null
            },
            Response = { Body = new MemoryStream() },
            RequestAborted = requestLifetime.Token
        };

        requestLifetime.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => middleware.InvokeAsync(context));
    }

    /// <summary>Cancels the request lifetime after middleware has finished reading its body.</summary>
    private sealed class CancelOnEndStream(byte[] bytes, CancellationTokenSource lifetime)
        : MemoryStream(bytes)
    {
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            var read = Read(buffer, offset, count);
            if (read == 0) lifetime.Cancel();
            return Task.FromResult(read);
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            var read = Read(buffer.Span);
            if (read == 0) lifetime.Cancel();
            return ValueTask.FromResult(read);
        }
    }

    /// <summary>A stream of a given length that remembers how much of it was ever read.</summary>
    private sealed class CountingStream(long length) : Stream
    {
        /// <summary>How many bytes readers have taken so far.</summary>
        public long BytesRead { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => length;

        public override long Position { get; set; }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var remaining = length - Position;
            if (remaining <= 0) return 0;
            var n = (int)Math.Min(count, remaining);
            Array.Fill(buffer, (byte)' ', offset, n);
            Position += n;
            BytesRead += n;
            return n;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return origin switch
            {
                SeekOrigin.Begin => Position = offset,
                SeekOrigin.Current => Position += offset,
                _ => Position = length + offset
            };
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }
    }
}
