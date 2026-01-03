using AsposeMcpServer.Core.Security;

namespace AsposeMcpServer.Tests.Core.Security;

public class AuthConfigTests
{
    [Fact]
    public void ApiKeyConfig_DefaultValues_ShouldBeCorrect()
    {
        var config = new ApiKeyConfig();

        Assert.False(config.Enabled);
        Assert.Equal(ApiKeyMode.Local, config.Mode);
        Assert.Equal("X-API-Key", config.HeaderName);
        Assert.Equal("X-Tenant-Id", config.TenantIdHeader);
        Assert.Equal(5, config.CustomTimeoutSeconds);
        Assert.Null(config.Keys);
    }

    [Fact]
    public void JwtConfig_DefaultValues_ShouldBeCorrect()
    {
        var config = new JwtConfig();

        Assert.False(config.Enabled);
        Assert.Equal(JwtMode.Local, config.Mode);
        Assert.Equal("tenant_id", config.TenantIdClaim);
        Assert.Equal("X-Tenant-Id", config.TenantIdHeader);
        Assert.Equal("X-User-Id", config.UserIdHeader);
        Assert.Equal(5, config.CustomTimeoutSeconds);
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_WithApiKeySettings()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_ENABLED", "true");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_MODE", "gateway");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_KEYS", "key1:tenant1,key2:tenant2");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.True(config.ApiKey.Enabled);
            Assert.Equal(ApiKeyMode.Gateway, config.ApiKey.Mode);
            Assert.NotNull(config.ApiKey.Keys);
            Assert.Equal(2, config.ApiKey.Keys.Count);
            Assert.Equal("tenant1", config.ApiKey.Keys["key1"]);
            Assert.Equal("tenant2", config.ApiKey.Keys["key2"]);
        }
        finally
        {
            // Cleanup
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_ENABLED", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_MODE", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_KEYS", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_WithJwtSettings()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_ENABLED", "true");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_MODE", "introspection");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_ISSUER", "test-issuer");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_AUDIENCE", "test-audience");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.True(config.Jwt.Enabled);
            Assert.Equal(JwtMode.Introspection, config.Jwt.Mode);
            Assert.Equal("test-issuer", config.Jwt.Issuer);
            Assert.Equal("test-audience", config.Jwt.Audience);
        }
        finally
        {
            // Cleanup
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_ENABLED", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_MODE", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_ISSUER", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_AUDIENCE", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_WithRateLimit()
    {
        Environment.SetEnvironmentVariable("ASPOSE_RATE_LIMIT", "100");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal(100, config.RateLimitPerMinute);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_RATE_LIMIT", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_WithAllowedOrigins()
    {
        Environment.SetEnvironmentVariable("ASPOSE_ALLOWED_ORIGINS", "http://localhost:3000,https://example.com");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.NotNull(config.AllowedOrigins);
            Assert.Equal(2, config.AllowedOrigins.Length);
            Assert.Contains("http://localhost:3000", config.AllowedOrigins);
            Assert.Contains("https://example.com", config.AllowedOrigins);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_ALLOWED_ORIGINS", null);
        }
    }
}