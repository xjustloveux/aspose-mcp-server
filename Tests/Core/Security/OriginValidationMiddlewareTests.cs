using AsposeMcpServer.Core.Security;
using Microsoft.AspNetCore.Http;

namespace AsposeMcpServer.Tests.Core.Security;

public class OriginValidationMiddlewareTests
{
    private static DefaultHttpContext CreateHttpContext(string? origin = null, string path = "/mcp")
    {
        var context = new DefaultHttpContext { Request = { Path = path } };

        if (origin != null) context.Request.Headers.Origin = origin;

        return context;
    }

    private static OriginValidationMiddleware CreateMiddleware(
        OriginValidationConfig config,
        RequestDelegate? next = null)
    {
        next ??= _ => Task.CompletedTask;
        return new OriginValidationMiddleware(next, config);
    }

    [Fact]
    public async Task InvokeAsync_WhenDisabled_ShouldPassThrough()
    {
        var config = new OriginValidationConfig { Enabled = false };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext("https://malicious.com");

        await middleware.InvokeAsync(context);

        Assert.True(nextCalled);
        Assert.NotEqual(StatusCodes.Status403Forbidden, context.Response.StatusCode);
    }

    [Fact]
    public async Task InvokeAsync_WithExcludedPath_ShouldPassThrough()
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            ExcludedPaths = ["/health"]
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext("https://malicious.com", "/health");

        await middleware.InvokeAsync(context);

        Assert.True(nextCalled);
    }

    [Fact]
    public async Task InvokeAsync_WithMissingOrigin_WhenAllowed_ShouldPassThrough()
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowMissingOrigin = true
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext();

        await middleware.InvokeAsync(context);

        Assert.True(nextCalled);
    }

    [Fact]
    public async Task InvokeAsync_WithMissingOrigin_WhenRequired_ShouldReject()
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowMissingOrigin = false
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext();

        await middleware.InvokeAsync(context);

        Assert.False(nextCalled);
        Assert.Equal(StatusCodes.Status403Forbidden, context.Response.StatusCode);
    }

    [Theory]
    [InlineData("http://localhost")]
    [InlineData("http://localhost:3000")]
    [InlineData("https://localhost")]
    [InlineData("http://127.0.0.1")]
    [InlineData("http://127.0.0.1:8080")]
    [InlineData("http://[::1]")]
    [InlineData("http://app.localhost")]
    [InlineData("http://sub.app.localhost")]
    public async Task InvokeAsync_WithLocalhostOrigin_WhenAllowed_ShouldPassThrough(string origin)
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowLocalhost = true
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext(origin);

        await middleware.InvokeAsync(context);

        Assert.True(nextCalled);
    }

    [Theory]
    [InlineData("http://localhost")]
    [InlineData("http://127.0.0.1")]
    public async Task InvokeAsync_WithLocalhostOrigin_WhenDisallowed_ShouldReject(string origin)
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowLocalhost = false,
            AllowMissingOrigin = false
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext(origin);

        await middleware.InvokeAsync(context);

        Assert.False(nextCalled);
        Assert.Equal(StatusCodes.Status403Forbidden, context.Response.StatusCode);
    }

    [Fact]
    public async Task InvokeAsync_WithAllowedOrigin_ShouldPassThrough()
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowLocalhost = false,
            AllowedOrigins = ["https://trusted.example.com"]
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext("https://trusted.example.com");

        await middleware.InvokeAsync(context);

        Assert.True(nextCalled);
    }

    [Fact]
    public async Task InvokeAsync_WithUnallowedOrigin_ShouldReject()
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowLocalhost = false,
            AllowedOrigins = ["https://trusted.example.com"]
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext("https://malicious.com");

        await middleware.InvokeAsync(context);

        Assert.False(nextCalled);
        Assert.Equal(StatusCodes.Status403Forbidden, context.Response.StatusCode);
    }

    [Fact]
    public async Task InvokeAsync_WithInvalidOriginUri_ShouldReject()
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowLocalhost = true,
            AllowMissingOrigin = false
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext("not-a-valid-uri");

        await middleware.InvokeAsync(context);

        Assert.False(nextCalled);
        Assert.Equal(StatusCodes.Status403Forbidden, context.Response.StatusCode);
    }

    [Fact]
    public async Task InvokeAsync_WithCaseInsensitiveOrigin_ShouldMatch()
    {
        var config = new OriginValidationConfig
        {
            Enabled = true,
            AllowLocalhost = false,
            AllowedOrigins = ["https://TRUSTED.EXAMPLE.COM"]
        };
        var nextCalled = false;
        var middleware = CreateMiddleware(config, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        var context = CreateHttpContext("https://trusted.example.com");

        await middleware.InvokeAsync(context);

        Assert.True(nextCalled);
    }

    [Fact]
    public void Constructor_WithNullNext_ShouldThrow()
    {
        var config = new OriginValidationConfig();

        Assert.Throws<ArgumentNullException>(() =>
            new OriginValidationMiddleware(null!, config));
    }

    [Fact]
    public void Constructor_WithNullConfig_ShouldThrow()
    {
        RequestDelegate next = _ => Task.CompletedTask;

        Assert.Throws<ArgumentNullException>(() =>
            new OriginValidationMiddleware(next, null!));
    }
}
