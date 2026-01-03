using AsposeMcpServer.Core.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace AsposeMcpServer.Tests.Core.Security;

public class ApiKeyAuthenticationMiddlewareTests
{
    private readonly ILogger<ApiKeyAuthenticationMiddleware> _logger =
        NullLogger<ApiKeyAuthenticationMiddleware>.Instance;

    [Fact]
    public async Task LocalMode_ValidKey_ShouldSetTenantId()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["valid-key"] = "tenant-123"
            }
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "valid-key";

        string? capturedTenantId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(
            ctx =>
            {
                capturedTenantId = ctx.Items["TenantId"]?.ToString();
                return Task.CompletedTask;
            },
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal("tenant-123", capturedTenantId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_InvalidKey_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["valid-key"] = "tenant-123"
            }
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "invalid-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_MissingKey_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["valid-key"] = "tenant-123"
            }
        };

        var context = CreateHttpContext();
        // No API key header

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task GatewayMode_WithTenantIdHeader_ShouldSetTenantId()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Gateway,
            HeaderName = "X-API-Key",
            TenantIdHeader = "X-Tenant-Id"
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "any-key";
        context.Request.Headers["X-Tenant-Id"] = "gateway-tenant";

        string? capturedTenantId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(
            ctx =>
            {
                capturedTenantId = ctx.Items["TenantId"]?.ToString();
                return Task.CompletedTask;
            },
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal("gateway-tenant", capturedTenantId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task GatewayMode_MissingTenantIdHeader_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Gateway,
            HeaderName = "X-API-Key",
            TenantIdHeader = "X-Tenant-Id"
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "any-key";
        // Missing X-Tenant-Id header

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task HealthEndpoint_ShouldSkipAuthentication()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            Keys = new Dictionary<string, string>()
        };

        var context = CreateHttpContext("/health");
        // No API key

        var nextCalled = false;
        var middleware = new ApiKeyAuthenticationMiddleware(
            _ =>
            {
                nextCalled = true;
                return Task.CompletedTask;
            },
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.True(nextCalled);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task MetricsEndpoint_ShouldSkipAuthentication()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            Keys = new Dictionary<string, string>()
        };

        var context = CreateHttpContext("/metrics");

        var nextCalled = false;
        var middleware = new ApiKeyAuthenticationMiddleware(
            _ =>
            {
                nextCalled = true;
                return Task.CompletedTask;
            },
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.True(nextCalled);
    }

    [Fact]
    public async Task LocalMode_NoKeysConfigured_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            Keys = null
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "any-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomHeaderName_ShouldReadFromCorrectHeader()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "Authorization-Key",
            Keys = new Dictionary<string, string>
            {
                ["my-api-key"] = "my-tenant"
            }
        };

        var context = CreateHttpContext();
        context.Request.Headers["Authorization-Key"] = "my-api-key";

        string? capturedTenantId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(
            ctx =>
            {
                capturedTenantId = ctx.Items["TenantId"]?.ToString();
                return Task.CompletedTask;
            },
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal("my-tenant", capturedTenantId);
    }

    private static DefaultHttpContext CreateHttpContext(string path = "/api/test")
    {
        var context = new DefaultHttpContext
        {
            Request = { Path = path, Method = "POST" },
            Response = { Body = new MemoryStream() }
        };
        return context;
    }
}