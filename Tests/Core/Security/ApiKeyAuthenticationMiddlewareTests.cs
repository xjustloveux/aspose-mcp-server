using System.Net;
using AsposeMcpServer.Core.Security;
using AsposeMcpServer.Core.Tracking;
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

    #region Credential Ownership

    /// <summary>
    ///     Runs one request through the middleware and reports the owner it was given.
    /// </summary>
    /// <param name="config">The configuration to run under.</param>
    /// <param name="key">The API key to present.</param>
    /// <returns>The group and user the session layer would see.</returns>
    private async Task<(string? GroupId, string? UserId)> OwnerFor(ApiKeyConfig config, string key)
    {
        var context = CreateHttpContext();
        context.Request.Headers[config.HeaderName] = key;

        string? groupId = null;
        string? userId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, ctx =>
        {
            groupId = ctx.Items["GroupId"]?.ToString();
            userId = ctx.Items["UserId"]?.ToString();
            return Task.CompletedTask;
        });

        return (groupId, userId);
    }

    /// <summary>
    ///     R4-S04 / R5-T02: two authenticated keys with no group of their own used to reach the
    ///     session layer looking exactly like unauthenticated requests, so they shared the anonymous
    ///     bucket and each could open the other's documents. An API key is stable for its whole life
    ///     — unlike a bearer token, which is reissued (R5-S02) — so a fingerprint of it is a bucket
    ///     that holds still across requests and across sessions.
    /// </summary>
    [Fact]
    public async Task LocalMode_AKeyWithNoGroup_ShouldGetTheSameOwnerOnEveryRequest()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["first-key"] = "",
                ["second-key"] = ""
            }
        };

        var first = await OwnerFor(config, "first-key");
        var again = await OwnerFor(config, "first-key");
        var other = await OwnerFor(config, "second-key");

        Assert.NotNull(first.UserId);
        Assert.StartsWith(CredentialOwnership.Prefix, first.UserId!, StringComparison.Ordinal);

        // The same key is the same owner: a caller's sessions stay reachable, and their per-owner
        // quota is not reset by asking again.
        Assert.Equal(first.UserId, again.UserId);

        // Different keys are different owners, which is what the shared anonymous bucket was not.
        Assert.NotEqual(first.UserId, other.UserId);

        // And the fingerprint is not the key.
        Assert.DoesNotContain("first-key", first.UserId!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task LocalMode_AKeyWithAGroup_ShouldKeepTheGroupAsItsOwner()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["grouped-key"] = "group-123"
            }
        };

        var owner = await OwnerFor(config, "grouped-key");

        Assert.Equal("group-123", owner.GroupId);
        Assert.Null(owner.UserId);
    }

    #endregion

    #region Local Mode Tests

    [Fact]
    public async Task LocalMode_ValidKey_ShouldSetGroupId()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            HeaderName = "X-API-Key",
            Keys = new Dictionary<string, string>
            {
                ["valid-key"] = "group-123"
            }
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "valid-key";

        string? capturedGroupId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            return Task.CompletedTask;
        });
        Assert.Equal("group-123", capturedGroupId);
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
                ["valid-key"] = "group-123"
            }
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "invalid-key";

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
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
                ["valid-key"] = "group-123"
            }
        };

        var context = CreateHttpContext();

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
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

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
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
                ["my-api-key"] = "my-group"
            }
        };

        var context = CreateHttpContext();
        context.Request.Headers["Authorization-Key"] = "my-api-key";

        string? capturedGroupId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            return Task.CompletedTask;
        });
        Assert.Equal("my-group", capturedGroupId);
    }

    #endregion

    #region Gateway Mode Tests

    [Fact]
    public async Task GatewayMode_WithGroupIdHeader_ShouldSetGroupId()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Gateway,
            // A-04: gateway mode trusts nobody until a proxy is named; these cases are about
            // header handling, so they declare the boundary as enforced elsewhere.
            TrustedProxies = [TrustedProxyEvaluator.TrustAnyEntry],
            HeaderName = "X-API-Key",
            GroupIdentifierHeader = "X-Group-Id"
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "any-key";
        context.Request.Headers["X-Group-Id"] = "gateway-group";

        string? capturedGroupId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            return Task.CompletedTask;
        });
        Assert.Equal("gateway-group", capturedGroupId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task GatewayMode_MissingGroupIdHeader_ShouldBeRejected()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Gateway,
            // A-04: gateway mode trusts nobody until a proxy is named; these cases are about
            // header handling, so they declare the boundary as enforced elsewhere.
            TrustedProxies = [TrustedProxyEvaluator.TrustAnyEntry],
            HeaderName = "X-API-Key",
            GroupIdentifierHeader = "X-Group-Id"
        };

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "any-key";

        var reachedTheServer = false;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ =>
        {
            reachedTheServer = true;
            return Task.CompletedTask;
        });

        // The group header is the whole identity on this path, so accepting a request without it
        // made every such request the same owner (R3-S02).
        Assert.False(reachedTheServer);
        Assert.Equal(401, context.Response.StatusCode);
    }

    #endregion

    #region Introspection Mode Tests

    [Fact]
    public async Task IntrospectionMode_ActiveResponse_ShouldSetGroupId()
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
            """{"active": true, "group_id": "introspection-group"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "some-api-key";

        string? capturedGroupId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            return Task.CompletedTask;
        });

        Assert.Equal("introspection-group", capturedGroupId);
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

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

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

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

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

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

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
                    Content = new StringContent("""{"active": true, "group_id": "test-group"}""")
                };
            });

        var httpClientFactory = CreateMockHttpClientFactory(mockHandler.Object);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "my-api-key";

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

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
                    Content = new StringContent("""{"active": true, "group_id": "test-group"}""")
                };
            });

        var httpClientFactory = CreateMockHttpClientFactory(mockHandler.Object);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "my-api-key";

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

        Assert.NotNull(capturedContent);
        Assert.Contains("key=my-api-key", capturedContent);
    }

    #endregion

    #region Custom Mode Tests

    [Fact]
    public async Task CustomMode_ValidResponse_ShouldSetGroupId()
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
            """{"valid": true, "group_id": "custom-group"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers["X-API-Key"] = "custom-key";

        string? capturedGroupId = null;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            return Task.CompletedTask;
        });

        Assert.Equal("custom-group", capturedGroupId);
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

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

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

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

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

        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

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
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        Assert.True(nextCalled);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task MetricsEndpoint_ShouldRequireAuthenticationByDefault()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            Keys = new Dictionary<string, string>()
        };

        var context = CreateHttpContext("/metrics");

        var nextCalled = false;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });

        Assert.False(nextCalled);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task MetricsEndpoint_ShouldSkipAuthentication_WhenAnonymousAccessConfigured()
    {
        var config = new ApiKeyConfig
        {
            Enabled = true,
            Mode = ApiKeyMode.Local,
            Keys = new Dictionary<string, string>()
        };
        var tracking = new TrackingConfig { MetricsEnabled = true, MetricsRequireAuth = false };

        var context = CreateHttpContext("/metrics");

        var nextCalled = false;
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger, null, tracking);
        await middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });

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
        var middleware = new ApiKeyAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
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
