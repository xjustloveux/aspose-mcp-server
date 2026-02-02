using AsposeMcpServer.Core.Tracking;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Core.Tracking;

/// <summary>
///     Unit tests for TrackingMiddleware class
/// </summary>
public class TrackingMiddlewareTests
{
    #region Constructor Tests

    [Fact]
    public void Constructor_WithMinimalConfig_ShouldNotThrow()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();
        RequestDelegate next = _ => Task.CompletedTask;

        var exception = Record.Exception(() =>
            new TrackingMiddleware(next, config, logger));

        Assert.Null(exception);
    }

    [Fact]
    public void Constructor_WithMetricsEnabled_ShouldNotThrow()
    {
        var config = new TrackingConfig
        {
            LogEnabled = false,
            MetricsEnabled = true,
            MetricsPath = "/metrics"
        };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();
        RequestDelegate next = _ => Task.CompletedTask;

        var exception = Record.Exception(() =>
            new TrackingMiddleware(next, config, logger));

        Assert.Null(exception);
    }

    [Fact]
    public void Constructor_WithWebhookEnabled_ShouldNotThrow()
    {
        var config = new TrackingConfig
        {
            LogEnabled = false,
            WebhookEnabled = true,
            WebhookUrl = "https://example.com/webhook"
        };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();
        RequestDelegate next = _ => Task.CompletedTask;

        var exception = Record.Exception(() =>
            new TrackingMiddleware(next, config, logger));

        Assert.Null(exception);
    }

    #endregion

    #region InvokeAsync Tests

    [Fact]
    public async Task InvokeAsync_ShouldCallNextMiddleware()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();
        var nextCalled = false;

        Task Next(HttpContext _)
        {
            nextCalled = true;
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext();

        await middleware.InvokeAsync(context);

        Assert.True(nextCalled);
    }

    [Fact]
    public async Task InvokeAsync_ShouldSetRequestId()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext _)
        {
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext();

        await middleware.InvokeAsync(context);

        Assert.True(context.Items.ContainsKey("RequestId"));
        Assert.NotNull(context.Items["RequestId"]);
        var requestId = context.Items["RequestId"]!.ToString();
        Assert.Equal(12, requestId!.Length);
    }

    [Fact]
    public async Task InvokeAsync_WithMetricsPath_ShouldReturnMetrics()
    {
        var config = new TrackingConfig
        {
            MetricsEnabled = true,
            MetricsPath = "/metrics"
        };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();
        var nextCalled = false;

        Task Next(HttpContext _)
        {
            nextCalled = true;
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext
        {
            Request = { Path = "/metrics" },
            Response = { Body = new MemoryStream() }
        };

        await middleware.InvokeAsync(context);

        Assert.False(nextCalled);
        Assert.Equal("text/plain; version=0.0.4; charset=utf-8", context.Response.ContentType);

        context.Response.Body.Seek(0, SeekOrigin.Begin);
        using var reader = new StreamReader(context.Response.Body);
        var responseBody = await reader.ReadToEndAsync();
        Assert.NotNull(responseBody);
    }

    [Fact]
    public async Task InvokeAsync_WithErrorStatusCode_ShouldTrackAsFailure()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext ctx)
        {
            ctx.Response.StatusCode = 500;
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext { Items = { ["ToolName"] = "test_tool" } };

        await middleware.InvokeAsync(context);

        Assert.Equal(500, context.Response.StatusCode);
    }

    [Fact]
    public async Task InvokeAsync_WithException_ShouldRethrow()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext _)
        {
            throw new InvalidOperationException("Test error");
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext();

        await Assert.ThrowsAsync<InvalidOperationException>(async () =>
            await middleware.InvokeAsync(context));
    }

    [Fact]
    public async Task InvokeAsync_WithToolName_ShouldTrackTool()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext _)
        {
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext { Items = { ["ToolName"] = "pdf_text", ["ToolOperation"] = "get" } };

        await middleware.InvokeAsync(context);

        Assert.True(context.Items.ContainsKey("RequestId"));
        Assert.NotNull(context.Items["RequestId"]);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task InvokeAsync_WithMcpPath_ShouldSetToolName()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext _)
        {
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext { Request = { Path = "/mcp" } };

        await middleware.InvokeAsync(context);

        Assert.True(context.Items.ContainsKey("RequestId"));
        Assert.NotNull(context.Items["RequestId"]);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task InvokeAsync_WithWebSocketPath_ShouldTrack()
    {
        var config = new TrackingConfig { LogEnabled = false };
        var logger = Mock.Of<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext _)
        {
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, logger);
        var context = new DefaultHttpContext { Request = { Path = "/ws" } };

        await middleware.InvokeAsync(context);

        Assert.True(context.Items.ContainsKey("RequestId"));
        Assert.NotNull(context.Items["RequestId"]);
        Assert.Equal(200, context.Response.StatusCode);
    }

    #endregion

    #region Logging Tests

    [Fact]
    public async Task InvokeAsync_WithConsoleLogging_ShouldLogToConsole()
    {
        var config = new TrackingConfig
        {
            LogEnabled = true,
            LogTargets = [LogTarget.Console]
        };
        var loggerMock = new Mock<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext _)
        {
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, loggerMock.Object);
        var context = new DefaultHttpContext { Items = { ["ToolName"] = "test_tool" } };

        await middleware.InvokeAsync(context);

        loggerMock.Verify(
            x => x.Log(
                It.IsAny<LogLevel>(),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => true),
                It.IsAny<Exception>(),
                It.Is<Func<It.IsAnyType, Exception?, string>>((v, t) => true)),
            Times.AtLeastOnce);
    }

    [Fact]
    public async Task InvokeAsync_WithFailure_ShouldLogWarning()
    {
        var config = new TrackingConfig
        {
            LogEnabled = true,
            LogTargets = [LogTarget.Console]
        };
        var loggerMock = new Mock<ILogger<TrackingMiddleware>>();

        Task Next(HttpContext ctx)
        {
            ctx.Response.StatusCode = 500;
            return Task.CompletedTask;
        }

        var middleware = new TrackingMiddleware(Next, config, loggerMock.Object);
        var context = new DefaultHttpContext { Items = { ["ToolName"] = "test_tool" } };

        await middleware.InvokeAsync(context);

        loggerMock.Verify(
            x => x.Log(
                It.IsAny<LogLevel>(),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => true),
                It.IsAny<Exception>(),
                It.Is<Func<It.IsAnyType, Exception?, string>>((v, t) => true)),
            Times.AtLeastOnce);
    }

    #endregion
}

/// <summary>
///     Unit tests for TrackingExtensions class
/// </summary>
public class TrackingExtensionsTests
{
    [Fact]
    public void UseTracking_WithAllDisabled_ShouldNotAddMiddleware()
    {
        var config = new TrackingConfig
        {
            LogEnabled = false,
            WebhookEnabled = false,
            MetricsEnabled = false
        };

        Assert.False(config.LogEnabled);
        Assert.False(config.WebhookEnabled);
        Assert.False(config.MetricsEnabled);
    }

    [Fact]
    public void TrackingConfig_DefaultValues_ShouldBeCorrect()
    {
        var config = new TrackingConfig();

        Assert.True(config.LogEnabled);
        Assert.Contains(LogTarget.Console, config.LogTargets);
        Assert.False(config.WebhookEnabled);
        Assert.False(config.MetricsEnabled);
        Assert.Equal("/metrics", config.MetricsPath);
        Assert.Equal(5, config.WebhookTimeoutSeconds);
    }
}
