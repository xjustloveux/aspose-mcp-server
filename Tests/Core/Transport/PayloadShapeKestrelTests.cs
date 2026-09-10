using System.Net;
using System.Net.Http.Headers;
using AsposeMcpServer.Core.Transport;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Logging;

namespace AsposeMcpServer.Tests.Core.Transport;

/// <summary>
///     R22-DEP02: the body cap on a real Kestrel listener, not a <c>DefaultHttpContext</c>.
///     <para>
///         The in-process fixture proved the middleware stops reading at the cap. A listener is a
///         different question: the host has buffers, a drain-on-keep-alive and a body-size feature
///         of its own between the wire and the middleware. Measured on loopback, on the port the
///         OS gives: of a body three times the cap, the listener took the cap plus about two
///         megabytes the client had already put on the wire. Evidence of the bound on a real
///         listener, not a guard that a single line's revert turns red: Kestrel's own
///         MaxRequestBodySize, set by the middleware, bounds the drain as well.
///     </para>
/// </summary>
public class PayloadShapeKestrelTests
{
    [Fact]
    public async Task ARealListener_ShouldRefuseAChunkedBodyOverTheCap_WithoutTakingItWhole()
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            ApplicationName = typeof(PayloadShapeKestrelTests).Assembly.GetName().Name
        });
        builder.Logging.ClearProviders();
        builder.WebHost.UseUrls("http://127.0.0.1:0");
        await using var app = builder.Build();

        var reached = false;
        app.UseMiddleware<PayloadShapeMiddleware>();
        app.Run(_ =>
        {
            reached = true;
            return Task.CompletedTask;
        });
        await app.StartAsync();

        var counting = new CountingStream(PayloadShapeMiddleware.MaxBodyBytes * 3);
        try
        {
            using var client = new HttpClient();
            // No length and no seeking: the client sends it chunked, as a streaming caller would.
            using var content = new StreamContent(counting);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            HttpStatusCode? status = null;
            try
            {
                using var response = await client.PostAsync(app.Urls.First() + "/mcp", content);
                status = response.StatusCode;
            }
            catch (HttpRequestException)
            {
                // The listener may close the connection before the client finishes sending; the
                // client then reports the reset rather than the response. What matters below is
                // measured on the stream, not on the response.
            }

            Assert.False(reached, "a body over the cap reached the endpoint");
            if (status != null) Assert.Equal(HttpStatusCode.RequestEntityTooLarge, status);
            // What the client handed over includes what the kernel buffered ahead of the server's
            // reads: about 2 MB past the cap on Windows loopback, about 11 MB on Linux. The bound
            // that holds on both is that the body was not taken whole.
            Assert.True(counting.BytesRead < counting.TotalLength,
                $"the listener took the whole {counting.BytesRead:N0}-byte body it was meant to stop reading at "
                + $"{PayloadShapeMiddleware.MaxBodyBytes:N0}");
            Assert.True(
                counting.BytesRead < 2 * PayloadShapeMiddleware.MaxBodyBytes +
                (OperatingSystem.IsWindows() ? 0 : 4 * PayloadShapeMiddleware.MaxBodyBytes / 2),
                $"the listener took {counting.BytesRead:N0} bytes of a body it was meant to stop reading at "
                + $"{PayloadShapeMiddleware.MaxBodyBytes:N0} (the cap plus what the client had already put on the wire)");
        }
        finally
        {
            await app.StopAsync();
        }
    }

    /// <summary>An unseekable stream of a given length that remembers how much of it was ever read.</summary>
    private sealed class CountingStream(long length) : Stream
    {
        private long _position;

        /// <summary>How long the body is.</summary>
        public long TotalLength => length;

        /// <summary>How many bytes readers have taken so far.</summary>
        public long BytesRead { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var remaining = length - _position;
            if (remaining <= 0) return 0;
            var n = (int)Math.Min(count, remaining);
            Array.Fill(buffer, (byte)' ', offset, n);
            _position += n;
            BytesRead += n;
            return n;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
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
