using Microsoft.AspNetCore.Http.Features;

namespace AsposeMcpServer.Core.Transport;

/// <summary>
///     Puts <see cref="PayloadShapeGuard" /> in front of the Streamable HTTP route.
///     <para>
///         The guard was wired into the WebSocket handler and nowhere else, so a peer on the HTTP
///         transport could send the same shape and have the SDK bind it unguarded (R20-RES02).
///         The WebSocket route also had a byte cap of this server's own; HTTP relied on the
///         host's default. Both are applied here, to the same values, so the two network-facing
///         routes make the same promise.
///     </para>
///     <para>
///         stdio is not behind this. Its peer is the local client process that launched the
///         server, which the server's threat model treats as its own; a bound there would be a
///         bound on the operator's own machine talking to itself.
///     </para>
/// </summary>
public sealed class PayloadShapeMiddleware
{
    /// <summary>The most a request body may hold, matching the WebSocket route.</summary>
    public const long MaxBodyBytes = 10L * 1024 * 1024;

    private readonly RequestDelegate _next;

    /// <summary>Creates the middleware.</summary>
    /// <param name="next">The rest of the pipeline.</param>
    public PayloadShapeMiddleware(RequestDelegate next)
    {
        _next = next;
    }

    /// <summary>Refuses a body that is too large or a shape this server will not bind.</summary>
    /// <param name="context">The request.</param>
    /// <returns>A task that completes when the request has been handled or refused.</returns>
    public async Task InvokeAsync(HttpContext context)
    {
        if (context.Request.ContentLength is 0 or null && !context.Request.Body.CanRead)
        {
            await _next(context);
            return;
        }

        if (context.Features.Get<IHttpMaxRequestBodySizeFeature>() is { IsReadOnly: false } size)
            size.MaxRequestBodySize = MaxBodyBytes;

        if (context.Request.ContentLength > MaxBodyBytes)
        {
            await RefuseAsync(context);
            return;
        }

        // Buffered so it can be read here and again by the SDK. The cap above bounds what is held.
        context.Request.EnableBuffering((int)MaxBodyBytes);

        // Bounded by this middleware itself, not only by the host: a body of unknown length was
        // copied whole and measured afterwards, so on a host without the size feature — or a
        // test context — the copy was the unbounded allocation this guard exists to prevent
        // (R21-RES04). One byte past the cap is enough to know, and is all that is read.
        using var buffer = new MemoryStream();
        var chunk = new byte[81920];
        while (buffer.Length <= MaxBodyBytes)
        {
            var read = await context.Request.Body.ReadAsync(chunk, context.RequestAborted);
            if (read == 0) break;
            buffer.Write(chunk, 0, read);
        }

        context.Request.Body.Position = 0;

        if (buffer.Length > MaxBodyBytes)
        {
            await RefuseAsync(context);
            return;
        }

        if (!PayloadShapeGuard.IsWithinBounds(buffer.GetBuffer().AsSpan(0, (int)buffer.Length),
                out var refusal))
        {
            context.Response.StatusCode = StatusCodes.Status400BadRequest;
            await context.Response.WriteAsync(refusal, context.RequestAborted);
            return;
        }

        await _next(context);
    }

    /// <summary>Refuses a body that is, or has proven to be, over the cap.</summary>
    /// <param name="context">The request.</param>
    /// <returns>A task that completes when the refusal has been written.</returns>
    /// <remarks>
    ///     With <c>Connection: close</c>: a server that keeps the connection alive after
    ///     refusing an unread body has to drain that body first, and Kestrel does — so the bytes
    ///     this middleware stopped reading were read anyway, by the host, right after it. Closing
    ///     is what makes "not read" true on a real listener (R22-DEP02, measured on loopback).
    /// </remarks>
    private static async Task RefuseAsync(HttpContext context)
    {
        context.Response.StatusCode = StatusCodes.Status413PayloadTooLarge;
        context.Response.Headers.Connection = "close";
        await context.Response.WriteAsync(
            $"The request body may not exceed {MaxBodyBytes:N0} bytes.",
            context.RequestAborted);
    }
}
