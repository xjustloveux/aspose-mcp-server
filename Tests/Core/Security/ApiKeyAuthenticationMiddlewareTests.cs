using System.Net;
using AsposeMcpServer.Core.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;
using Moq.Protected;

namespace AsposeMcpServer.Tests.Core.Security;

/// <summary>
///     Unit tests for ApiKeyAuthenticationMiddleware class
/// </summary>
public class ApiKeyAuthenticationMiddlewareTests
{
    private readonly ILogger<ApiKeyAuthenticationMiddleware> _logger =
        NullLogger<ApiKeyAuthenticationMiddleware>.Instance;

    #region Local Mode Tests

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

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal(401, context.Response.StatusCode);
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
    public async Task LocalMode_CustomHeaderName_ShouldReadFromCorrectHeader()
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

    #endregion

    #region Gateway Mode Tests

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

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);
        Assert.Equal(401, context.Response.StatusCode);
    }

    #endregion

    #region Introspection Mode Tests

    [Fact]
    public async Task IntrospectionMode_ActiveResponse_ShouldSetTenantId()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/introspect",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"active": true, "tenant_id": "introspection-tenant"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "some-api-key";

        string? capturedTenantId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(
            ctx =>
            {
                capturedTenantId = ctx.Items["TenantId"]?.ToString();
                return Task.CompletedTask;
            },
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.Equal("introspection-tenant", capturedTenantId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task IntrospectionMode_InactiveResponse_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/introspect",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"active": false}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "invalid-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task IntrospectionMode_ServerError_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/introspect",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(HttpStatusCode.InternalServerError, "");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "some-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task IntrospectionMode_NoEndpointConfigured_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Introspection,
            IntrospectionEndpoint = null
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "some-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task IntrospectionMode_CustomKeyField_ShouldUseConfiguredFieldName()
    {
        string? capturedContent = null;
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/introspect",
            IntrospectionKeyField = "api_token",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = new Mock<HttpMessageHandler>();
        mockHandler.Protected()
            .Setup<Task<HttpResponseMessage>>(
                "SendAsync",
                ItExpr.IsAny<HttpRequestMessage>(),
                ItExpr.IsAny<CancellationToken>())
            .Returns<HttpRequestMessage, CancellationToken>(async (req, _) =>
            {
                capturedContent = await req.Content!.ReadAsStringAsync();
                return new HttpResponseMessage
                {
                    StatusCode = HttpStatusCode.OK,
                    Content = new StringContent("""{"active": true, "tenant_id": "test-tenant"}""")
                };
            });

        var httpClientFactory = CreateMockHttpClientFactory(mockHandler.Object);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "my-api-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.NotNull(capturedContent);
        Assert.Contains("api_token=my-api-key", capturedContent);
    }

    [Fact]
    public async Task IntrospectionMode_DefaultKeyField_ShouldUseKey()
    {
        string? capturedContent = null;
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/introspect",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = new Mock<HttpMessageHandler>();
        mockHandler.Protected()
            .Setup<Task<HttpResponseMessage>>(
                "SendAsync",
                ItExpr.IsAny<HttpRequestMessage>(),
                ItExpr.IsAny<CancellationToken>())
            .Returns<HttpRequestMessage, CancellationToken>(async (req, _) =>
            {
                capturedContent = await req.Content!.ReadAsStringAsync();
                return new HttpResponseMessage
                {
                    StatusCode = HttpStatusCode.OK,
                    Content = new StringContent("""{"active": true, "tenant_id": "test-tenant"}""")
                };
            });

        var httpClientFactory = CreateMockHttpClientFactory(mockHandler.Object);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "my-api-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.NotNull(capturedContent);
        Assert.Contains("key=my-api-key", capturedContent);
    }

    #endregion

    #region Custom Mode Tests

    [Fact]
    public async Task CustomMode_ValidResponse_ShouldSetTenantId()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Custom,
            CustomEndpoint = "https://auth.example.com/custom",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"valid": true, "tenant_id": "custom-tenant"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "custom-key";

        string? capturedTenantId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(
            ctx =>
            {
                capturedTenantId = ctx.Items["TenantId"]?.ToString();
                return Task.CompletedTask;
            },
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.Equal("custom-tenant", capturedTenantId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomMode_InvalidResponse_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Custom,
            CustomEndpoint = "https://auth.example.com/custom",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"valid": false, "error": "Invalid API key"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "invalid-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomMode_ServerError_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Custom,
            CustomEndpoint = "https://auth.example.com/custom",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(HttpStatusCode.ServiceUnavailable, "");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "some-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger,
            httpClientFactory);
        await middleware.InvokeAsync(context);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomMode_NoEndpointConfigured_ShouldReturn401()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Custom,
            CustomEndpoint = null
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "some-key";

        var middleware = new ApiKeyAuthenticationMiddleware(
            _ => Task.CompletedTask,
            config,
            _logger);
        await middleware.InvokeAsync(context);

        Assert.Equal(401, context.Response.StatusCode);
    }

    #endregion

    #region Endpoint Exclusion Tests

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
    public async Task ReadyEndpoint_ShouldSkipAuthentication()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            Keys = new Dictionary<string, string>()
        };

        var context = CreateHttpContext("/ready");

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

    #endregion

    #region Helper Methods

    private static DefaultHttpContext CreateHttpContext(string path = "/api/test")
    {
        var context = new DefaultHttpContext
        {
            Request = { Path = path, Method = "POST" },
            Response = { Body = new MemoryStream() }
        };
        return context;
    }

    private static HttpMessageHandler CreateMockHttpHandler(HttpStatusCode statusCode, string content)
    {
        var mockHandler = new Mock<HttpMessageHandler>();
        mockHandler.Protected()
            .Setup<Task<HttpResponseMessage>>(
                "SendAsync",
                ItExpr.IsAny<HttpRequestMessage>(),
                ItExpr.IsAny<CancellationToken>())
            .ReturnsAsync(new HttpResponseMessage
            {
                StatusCode = statusCode,
                Content = new StringContent(content)
            });
        return mockHandler.Object;
    }

    private static IHttpClientFactory CreateMockHttpClientFactory(HttpMessageHandler handler)
    {
        var mockFactory = new Mock<IHttpClientFactory>();
        mockFactory.Setup(f => f.CreateClient(It.IsAny<string>()))
            .Returns(new HttpClient(handler));
        return mockFactory.Object;
    }

    #endregion
}