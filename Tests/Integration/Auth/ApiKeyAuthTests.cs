using AsposeMcpServer.Core.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Integration.Auth;

/// <summary>
///     Integration tests for API Key authentication.
/// </summary>
[Trait("Category", "Integration")]
public class ApiKeyAuthTests : IDisposable
{
    private readonly ApiKeyAuthenticationMiddleware _middleware;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ApiKeyAuthTests" /> class.
    /// </summary>
    public ApiKeyAuthTests()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["valid-api-key-1"] = "group-a",
                ["valid-api-key-2"] = "group-b"
            }
        };

        var logger = Mock.Of<ILogger<ApiKeyAuthenticationMiddleware>>();
        _middleware = new ApiKeyAuthenticationMiddleware(config, logger);
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public void Dispose()
    {
        _middleware.Dispose();
    }

    /// <summary>
    ///     Verifies that a valid API key is authenticated successfully.
    /// </summary>
    [Fact]
    public async Task ApiKey_ValidKey_Authenticated()
    {
        // Arrange
        var context = CreateHttpContext("/mcp");
        context.Request.Headers["X-API-Key"] = "valid-api-key-1";
        var nextCalled = false;

        // Act
        await _middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });

        // Assert
        Assert.True(nextCalled);
        Assert.Equal("group-a", context.Items["GroupId"]);
    }

    /// <summary>
    ///     Verifies that an invalid API key is rejected.
    /// </summary>
    [Fact]
    public async Task ApiKey_InvalidKey_Rejected()
    {
        // Arrange
        var context = CreateHttpContext("/mcp");
        context.Request.Headers["X-API-Key"] = "invalid-api-key";
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
    ///     Verifies that a missing API key is rejected.
    /// </summary>
    [Fact]
    public async Task ApiKey_MissingKey_Rejected()
    {
        // Arrange
        var context = CreateHttpContext("/mcp");
        // No API key header set
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
    ///     Verifies that different valid keys map to different groups.
    /// </summary>
    [Fact]
    public async Task ApiKey_DifferentKeys_DifferentGroups()
    {
        // Arrange - Key 1
        var context1 = CreateHttpContext("/mcp");
        context1.Request.Headers["X-API-Key"] = "valid-api-key-1";

        // Act
        await _middleware.InvokeAsync(context1, _ => Task.CompletedTask);

        // Assert
        Assert.Equal("group-a", context1.Items["GroupId"]);

        // Arrange - Key 2
        var context2 = CreateHttpContext("/mcp");
        context2.Request.Headers["X-API-Key"] = "valid-api-key-2";

        // Act
        await _middleware.InvokeAsync(context2, _ => Task.CompletedTask);

        // Assert
        Assert.Equal("group-b", context2.Items["GroupId"]);
    }

    /// <summary>
    ///     Verifies that health endpoints don't require authentication.
    /// </summary>
    [Theory]
    [InlineData("/health")]
    [InlineData("/ready")]
    [InlineData("/metrics")]
    public async Task ApiKey_HealthEndpoints_NoAuthRequired(string path)
    {
        // Arrange
        var context = CreateHttpContext(path);
        // No API key header set
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
