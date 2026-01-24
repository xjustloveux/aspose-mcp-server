using AsposeMcpServer.Core.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Integration.Auth;

/// <summary>
///     Integration tests for authentication middleware behavior.
/// </summary>
[Trait("Category", "Integration")]
public class AuthMiddlewareTests : IDisposable
{
    private ApiKeyAuthenticationMiddleware? _middleware;

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public void Dispose()
    {
        _middleware?.Dispose();
    }

    /// <summary>
    ///     Verifies that public endpoints (health/ready/metrics) don't require authentication.
    /// </summary>
    [Fact]
    public async Task Middleware_PublicEndpoint_NoAuthRequired()
    {
        // Arrange
        var config = CreateEnabledConfig();
        _middleware = CreateMiddleware(config);

        var context = CreateHttpContext("/health");
        var nextCalled = false;

        // Act
        await _middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });

        // Assert
        Assert.True(nextCalled);
    }

    /// <summary>
    ///     Verifies that protected endpoints require authentication.
    /// </summary>
    [Fact]
    public async Task Middleware_ProtectedEndpoint_RequiresAuth()
    {
        // Arrange
        var config = CreateEnabledConfig();
        _middleware = CreateMiddleware(config);

        var context = CreateHttpContext("/mcp");
        // No API key set
        var nextCalled = false;

        // Act
        await _middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });

        // Assert
        Assert.False(nextCalled);
        Assert.Equal(StatusCodes.Status401Unauthorized, context.Response.StatusCode);
    }

    /// <summary>
    ///     Verifies that Gateway mode trusts incoming requests.
    /// </summary>
    [Fact]
    public async Task Middleware_GatewayMode_TrustsRequest()
    {
        // Arrange
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Gateway,
            HeaderName = "X-API-Key",
            GroupIdentifierHeader = "X-Group-Id"
        };
        _middleware = CreateMiddleware(config);

        var context = CreateHttpContext("/mcp");
        context.Request.Headers["X-API-Key"] = "any-key"; // Gateway mode trusts the key
        context.Request.Headers["X-Group-Id"] = "gateway-group";
        var nextCalled = false;

        // Act
        await _middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });

        // Assert
        Assert.True(nextCalled);
        Assert.Equal("gateway-group", context.Items["GroupId"]);
    }

    /// <summary>
    ///     Verifies that missing keys config in Local mode returns error.
    /// </summary>
    [Fact]
    public async Task Middleware_LocalModeNoKeys_ReturnsError()
    {
        // Arrange
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = null // No keys configured
        };
        _middleware = CreateMiddleware(config);

        var context = CreateHttpContext("/mcp");
        context.Request.Headers["X-API-Key"] = "some-key";
        var nextCalled = false;

        // Act
        await _middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });

        // Assert
        Assert.False(nextCalled);
        Assert.Equal(StatusCodes.Status401Unauthorized, context.Response.StatusCode);
    }

    private static ApiKeyConfig CreateEnabledConfig()
    {
        return new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["test-key"] = "test-group"
            }
        };
    }

    private static ApiKeyAuthenticationMiddleware CreateMiddleware(ApiKeyConfig config)
    {
        var logger = Mock.Of<ILogger<ApiKeyAuthenticationMiddleware>>();
        return new ApiKeyAuthenticationMiddleware(config, logger);
    }

    private static DefaultHttpContext CreateHttpContext(string path)
    {
        var context = new DefaultHttpContext
        {
            Request = { Path = path },
            Response = { Body = new MemoryStream() }
        };
        return context;
    }
}
