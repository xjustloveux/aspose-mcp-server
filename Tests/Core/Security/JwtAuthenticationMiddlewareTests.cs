using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Security.Claims;
using System.Text;
using AsposeMcpServer.Core.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.IdentityModel.Tokens;
using Moq;
using Moq.Protected;

namespace AsposeMcpServer.Tests.Core.Security;

public class JwtAuthenticationMiddlewareTests
{
    private const string TestSecret = "this-is-a-test-secret-key-with-sufficient-length-for-hs256";
    private readonly ILogger<JwtAuthenticationMiddleware> _logger = NullLogger<JwtAuthenticationMiddleware>.Instance;

    [Fact]
    public async Task LocalMode_ValidToken_ShouldSetGroupAndUserId()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret,
            GroupIdentifierClaim = "tenant_id"
        };

        var token = GenerateTestToken(new Dictionary<string, string>
        {
            ["tenant_id"] = "test-group",
            ["sub"] = "user-123"
        });

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = $"Bearer {token}";

        string? capturedGroupId = null;
        string? capturedUserId = null;
        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            capturedUserId = ctx.Items["UserId"]?.ToString();
            return Task.CompletedTask;
        });
        Assert.Equal("test-group", capturedGroupId);
        Assert.Equal("user-123", capturedUserId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_InvalidToken_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret
        };

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer invalid-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_ExpiredToken_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret
        };

        var token = GenerateTestToken(
            new Dictionary<string, string> { ["sub"] = "user" },
            DateTime.UtcNow.AddMinutes(-10));

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = $"Bearer {token}";

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_MissingAuthorizationHeader_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret
        };

        var context = CreateHttpContext();

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_NonBearerToken_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret
        };

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Basic dXNlcjpwYXNz";

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task GatewayMode_WithHeaders_ShouldSetGroupAndUserId()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Gateway,
            // A-04: gateway mode trusts nobody until a proxy is named; these cases are about
            // header handling, so they declare the boundary as enforced elsewhere.
            TrustedProxies = [TrustedProxyEvaluator.TrustAnyEntry],
            GroupIdentifierHeader = "X-Group-Id",
            UserIdHeader = "X-User-Id"
        };

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer any-token";
        context.Request.Headers["X-Group-Id"] = "gateway-group";
        context.Request.Headers["X-User-Id"] = "gateway-user";

        string? capturedGroupId = null;
        string? capturedUserId = null;
        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            capturedUserId = ctx.Items["UserId"]?.ToString();
            return Task.CompletedTask;
        });
        Assert.Equal("gateway-group", capturedGroupId);
        Assert.Equal("gateway-user", capturedUserId);
    }

    [Fact]
    public async Task GatewayMode_MissingHeaders_ShouldBeRejected()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Gateway,
            // A-04: gateway mode trusts nobody until a proxy is named; these cases are about
            // header handling, so they declare the boundary as enforced elsewhere.
            TrustedProxies = [TrustedProxyEvaluator.TrustAnyEntry],
            GroupIdentifierHeader = "X-Group-Id",
            UserIdHeader = "X-User-Id"
        };

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer any-token";

        var reachedTheServer = false;
        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ =>
        {
            reachedTheServer = true;
            return Task.CompletedTask;
        });

        // This case used to be named "ShouldAllowAsAnonymous" and asserted exactly that: an
        // authenticated gateway request naming neither group nor user became an anonymous
        // identity. Every such request became the *same* anonymous identity, so two unrelated
        // principals shared one session bucket (R3-S02).
        Assert.False(reachedTheServer);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task HealthEndpoint_ShouldSkipAuthentication()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret
        };

        var context = CreateHttpContext("/health");

        var nextCalled = false;
        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        Assert.True(nextCalled);
    }

    [Fact]
    public async Task LocalMode_WithIssuerValidation_ValidIssuer_ShouldPass()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret,
            Issuer = "test-issuer"
        };

        var token = GenerateTestToken(
            new Dictionary<string, string> { ["sub"] = "user" },
            issuer: "test-issuer");

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = $"Bearer {token}";

        var nextCalled = false;
        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ =>
        {
            nextCalled = true;
            return Task.CompletedTask;
        });
        Assert.True(nextCalled);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_WithIssuerValidation_InvalidIssuer_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret,
            Issuer = "expected-issuer"
        };

        var token = GenerateTestToken(
            new Dictionary<string, string> { ["sub"] = "user" },
            issuer: "wrong-issuer");

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = $"Bearer {token}";

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task LocalMode_NoSecretConfigured_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = null
        };

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer some-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);
        Assert.Equal(401, context.Response.StatusCode);
    }

    /// <summary>
    ///     Runs one request through the middleware and reports the owner it was given.
    /// </summary>
    /// <param name="config">The JWT configuration to run under.</param>
    /// <param name="token">The bearer token to present.</param>
    /// <returns>The group and user the session layer would see, and the response status.</returns>
    private async Task<(string? GroupId, string? UserId, int Status)> OwnerFor(JwtConfig config, string token)
    {
        var context = CreateHttpContext();
        context.Request.Headers.Authorization = $"Bearer {token}";

        string? groupId = null;
        string? userId = null;
        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, ctx =>
        {
            groupId = ctx.Items["GroupId"]?.ToString();
            userId = ctx.Items["UserId"]?.ToString();
            return Task.CompletedTask;
        });

        return (groupId, userId, context.Response.StatusCode);
    }

    /// <summary>
    ///     R5-S02: a token carrying no configured claim used to be bucketed by a fingerprint of its
    ///     own bytes. Those change on every reissue, so the same principal landed somewhere new
    ///     each time — their earlier sessions became unreachable, and the per-owner session quota
    ///     reset as often as they asked for a fresh token. The subject is what names them, and it
    ///     survives the reissue.
    /// </summary>
    [Fact]
    public async Task LocalMode_TwoTokensForOnePrincipal_ShouldLandOnOneOwner()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret,
            GroupIdentifierClaim = "tenant_id",
            UserIdClaim = "preferred_username"
        };

        var first = await OwnerFor(config, GenerateTestToken(new Dictionary<string, string>
        {
            ["sub"] = "principal-1",
            ["jti"] = "token-one"
        }, DateTime.UtcNow.AddHours(1)));

        var reissued = await OwnerFor(config, GenerateTestToken(new Dictionary<string, string>
        {
            ["sub"] = "principal-1",
            ["jti"] = "token-two"
        }, DateTime.UtcNow.AddHours(2)));

        Assert.Equal(200, first.Status);
        Assert.Equal(200, reissued.Status);
        Assert.Equal("principal-1", first.UserId);
        Assert.Equal(first.UserId, reissued.UserId);
    }

    [Fact]
    public async Task LocalMode_TokensForDifferentPrincipals_ShouldStayApart()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret,
            GroupIdentifierClaim = "tenant_id",
            UserIdClaim = "preferred_username"
        };

        var one = await OwnerFor(config, GenerateTestToken(new Dictionary<string, string>
        {
            ["sub"] = "principal-1"
        }));
        var other = await OwnerFor(config, GenerateTestToken(new Dictionary<string, string>
        {
            ["client_id"] = "principal-2"
        }));

        Assert.Equal("principal-1", one.UserId);
        Assert.Equal("principal-2", other.UserId);
        Assert.NotEqual(one.UserId, other.UserId);
    }

    /// <summary>
    ///     A token naming nobody cannot be isolated from any other, so it is refused rather than
    ///     given a bucket that moves under it or one it shares with strangers.
    /// </summary>
    [Fact]
    public async Task LocalMode_TokenThatNamesNobody_ShouldBeRefused()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Local,
            Secret = TestSecret,
            GroupIdentifierClaim = "tenant_id",
            UserIdClaim = "preferred_username"
        };

        var refused = await OwnerFor(config, GenerateTestToken(new Dictionary<string, string>
        {
            ["scope"] = "documents:write"
        }));

        Assert.Equal(401, refused.Status);
        Assert.Null(refused.UserId);
        Assert.Null(refused.GroupId);
    }

    private static string GenerateTestToken(
        Dictionary<string, string> claims,
        DateTime? expires = null,
        string issuer = "test")
    {
        var securityKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(TestSecret));
        var credentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256);

        var claimsList = claims.Select(c => new Claim(c.Key, c.Value)).ToList();

        var token = new JwtSecurityToken(
            issuer,
            "test-audience",
            claimsList,
            expires: expires ?? DateTime.UtcNow.AddHours(1),
            signingCredentials: credentials);

        return new JwtSecurityTokenHandler().WriteToken(token);
    }

    [Fact]
    public async Task IntrospectionMode_ActiveResponse_ShouldSetGroupAndUserId()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/oauth/introspect",
            ClientId = "test-client",
            ClientSecret = "test-secret",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"active": true, "group_id": "introspection-group", "sub": "introspection-user"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer some-token";

        string? capturedGroupId = null;
        string? capturedUserId = null;
        var middleware = new JwtAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            capturedUserId = ctx.Items["UserId"]?.ToString();
            return Task.CompletedTask;
        });

        Assert.Equal("introspection-group", capturedGroupId);
        Assert.Equal("introspection-user", capturedUserId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task IntrospectionMode_InactiveResponse_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/oauth/introspect",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"active": false}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer invalid-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task IntrospectionMode_ServerError_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Introspection,
            IntrospectionEndpoint = "https://auth.example.com/oauth/introspect",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(HttpStatusCode.InternalServerError, "");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer some-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task IntrospectionMode_NoEndpointConfigured_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Introspection,
            IntrospectionEndpoint = null
        };

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer some-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomMode_ValidResponse_ShouldSetGroupAndUserId()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Custom,
            CustomEndpoint = "https://auth.example.com/custom",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"valid": true, "group_id": "custom-group", "user_id": "custom-user"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer custom-token";

        string? capturedGroupId = null;
        string? capturedUserId = null;
        var middleware = new JwtAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, ctx =>
        {
            capturedGroupId = ctx.Items["GroupId"]?.ToString();
            capturedUserId = ctx.Items["UserId"]?.ToString();
            return Task.CompletedTask;
        });

        Assert.Equal("custom-group", capturedGroupId);
        Assert.Equal("custom-user", capturedUserId);
        Assert.Equal(200, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomMode_InvalidResponse_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Custom,
            CustomEndpoint = "https://auth.example.com/custom",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(
            HttpStatusCode.OK,
            """{"valid": false, "error": "Token expired"}""");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer expired-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomMode_ServerError_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Custom,
            CustomEndpoint = "https://auth.example.com/custom",
            ExternalTimeoutSeconds = 5
        };

        var mockHandler = CreateMockHttpHandler(HttpStatusCode.ServiceUnavailable, "");
        var httpClientFactory = CreateMockHttpClientFactory(mockHandler);

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer some-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger, httpClientFactory);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

        Assert.Equal(401, context.Response.StatusCode);
    }

    [Fact]
    public async Task CustomMode_NoEndpointConfigured_ShouldReturn401()
    {
        var config = new JwtConfig
        {
            Enabled = true,
            Mode = JwtMode.Custom,
            CustomEndpoint = null
        };

        var context = CreateHttpContext();
        context.Request.Headers.Authorization = "Bearer some-token";

        var middleware = new JwtAuthenticationMiddleware(config, _logger);
        await middleware.InvokeAsync(context, _ => Task.CompletedTask);

        Assert.Equal(401, context.Response.StatusCode);
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
}
