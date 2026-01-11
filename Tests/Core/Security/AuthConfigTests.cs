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
        Assert.Equal("X-Group-Id", config.GroupIdentifierHeader);
        Assert.Equal(5, config.ExternalTimeoutSeconds);
        Assert.Equal("key", config.IntrospectionKeyField);
        Assert.Null(config.Keys);
    }

    [Fact]
    public void JwtConfig_DefaultValues_ShouldBeCorrect()
    {
        var config = new JwtConfig();

        Assert.False(config.Enabled);
        Assert.Equal(JwtMode.Local, config.Mode);
        Assert.Equal("tenant_id", config.GroupIdentifierClaim);
        Assert.Equal("sub", config.UserIdClaim);
        Assert.Equal("X-Group-Id", config.GroupIdentifierHeader);
        Assert.Equal("X-User-Id", config.UserIdHeader);
        Assert.Equal(5, config.ExternalTimeoutSeconds);
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
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_ENABLED", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_MODE", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_ISSUER", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_AUDIENCE", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyWithColonInTenant()
    {
        // Tenant IDs may contain colons - only the first colon should be used as separator
        // Format: key:tenant where tenant can contain colons
        var args = new[] { "--auth-apikey-keys:my-key:tenant:with:colons,simple-key:tenant2" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.NotNull(config.ApiKey.Keys);
        Assert.Equal(2, config.ApiKey.Keys.Count);
        Assert.Equal("tenant:with:colons", config.ApiKey.Keys["my-key"]);
        Assert.Equal("tenant2", config.ApiKey.Keys["simple-key"]);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyIntrospectionUrl()
    {
        var args = new[] { "--auth-apikey-introspection-url:https://auth.example.com/introspect" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal("https://auth.example.com/introspect", config.ApiKey.IntrospectionEndpoint);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyIntrospectionField()
    {
        var args = new[] { "--auth-apikey-introspection-field:api_token" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal("api_token", config.ApiKey.IntrospectionKeyField);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyTimeout()
    {
        var args = new[] { "--auth-apikey-timeout:10" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal(10, config.ApiKey.ExternalTimeoutSeconds);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtTimeout()
    {
        var args = new[] { "--auth-jwt-timeout:15" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal(15, config.Jwt.ExternalTimeoutSeconds);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyCustomUrl()
    {
        var args = new[] { "--auth-apikey-custom-url=https://custom.example.com/validate" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal("https://custom.example.com/validate", config.ApiKey.CustomEndpoint);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtIntrospectionSettings()
    {
        var args = new[]
        {
            "--auth-jwt-introspection-url:https://auth.example.com/oauth/introspect",
            "--auth-jwt-client-id:my-client",
            "--auth-jwt-client-secret:my-secret"
        };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal("https://auth.example.com/oauth/introspect", config.Jwt.IntrospectionEndpoint);
        Assert.Equal("my-client", config.Jwt.ClientId);
        Assert.Equal("my-secret", config.Jwt.ClientSecret);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtPublicKeyPath()
    {
        var args = new[] { "--auth-jwt-public-key-path=/path/to/public.pem" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal("/path/to/public.pem", config.Jwt.PublicKeyPath);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtCustomUrl()
    {
        var args = new[] { "--auth-jwt-custom-url=https://custom.example.com/validate-token" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal("https://custom.example.com/validate-token", config.Jwt.CustomEndpoint);
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_ApiKeyWithColonInKey()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_KEYS", "base64:encoded:key:tenant1");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.NotNull(config.ApiKey.Keys);
            Assert.Single(config.ApiKey.Keys);
            // First colon separates key from tenant, rest belongs to tenant
            Assert.Equal("encoded:key:tenant1", config.ApiKey.Keys["base64"]);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_KEYS", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtUserClaim()
    {
        var args = new[] { "--auth-jwt-user-claim:user_id" };

        var config = AuthConfig.LoadFromArgs(args);

        Assert.Equal("user_id", config.Jwt.UserIdClaim);
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_JwtUserClaim()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_CLAIM", "custom_user_claim");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal("custom_user_claim", config.Jwt.UserIdClaim);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_CLAIM", null);
        }
    }
}
