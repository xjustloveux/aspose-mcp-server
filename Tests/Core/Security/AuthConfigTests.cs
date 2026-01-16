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

    #region LoadFromEnvironment Additional Tests

    [Fact]
    public void AuthConfig_LoadFromEnvironment_ApiKeyHeaderAndGroupHeader()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_HEADER", "X-Custom-API-Key");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_GROUP_HEADER", "X-Custom-Group");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal("X-Custom-API-Key", config.ApiKey.HeaderName);
            Assert.Equal("X-Custom-Group", config.ApiKey.GroupIdentifierHeader);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_HEADER", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_GROUP_HEADER", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_ApiKeyIntrospectionSettings()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_URL", "https://introspect.example.com");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_AUTH", "Bearer token123");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CUSTOM_URL", "https://custom.example.com");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_FIELD", "api_token");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal("https://introspect.example.com", config.ApiKey.IntrospectionEndpoint);
            Assert.Equal("Bearer token123", config.ApiKey.IntrospectionAuthHeader);
            Assert.Equal("https://custom.example.com", config.ApiKey.CustomEndpoint);
            Assert.Equal("api_token", config.ApiKey.IntrospectionKeyField);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_URL", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_AUTH", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CUSTOM_URL", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_FIELD", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_ApiKeyTimeoutAndCache()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_TIMEOUT", "30");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_ENABLED", "true");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_TTL", "600");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_MAX_SIZE", "5000");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal(30, config.ApiKey.ExternalTimeoutSeconds);
            Assert.True(config.ApiKey.CacheEnabled);
            Assert.Equal(600, config.ApiKey.CacheTtlSeconds);
            Assert.Equal(5000, config.ApiKey.CacheMaxSize);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_TIMEOUT", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_ENABLED", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_TTL", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_MAX_SIZE", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_JwtSecretAndPublicKey()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_SECRET", "my-secret-key");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_PUBLIC_KEY_PATH", "/path/to/key.pem");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal("my-secret-key", config.Jwt.Secret);
            Assert.Equal("/path/to/key.pem", config.Jwt.PublicKeyPath);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_SECRET", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_PUBLIC_KEY_PATH", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_JwtGroupClaimAndHeaders()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_GROUP_CLAIM", "organization_id");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_GROUP_HEADER", "X-Org-Id");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_HEADER", "X-User");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal("organization_id", config.Jwt.GroupIdentifierClaim);
            Assert.Equal("X-Org-Id", config.Jwt.GroupIdentifierHeader);
            Assert.Equal("X-User", config.Jwt.UserIdHeader);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_GROUP_CLAIM", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_GROUP_HEADER", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_HEADER", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_JwtIntrospectionAndCustom()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_INTROSPECTION_URL", "https://oauth.example.com/introspect");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_ID", "client-123");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_SECRET", "secret-456");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CUSTOM_URL", "https://custom.example.com/jwt");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal("https://oauth.example.com/introspect", config.Jwt.IntrospectionEndpoint);
            Assert.Equal("client-123", config.Jwt.ClientId);
            Assert.Equal("secret-456", config.Jwt.ClientSecret);
            Assert.Equal("https://custom.example.com/jwt", config.Jwt.CustomEndpoint);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_INTROSPECTION_URL", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_ID", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_SECRET", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CUSTOM_URL", null);
        }
    }

    [Fact]
    public void AuthConfig_LoadFromEnvironment_JwtTimeoutAndCache()
    {
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_TIMEOUT", "20");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_ENABLED", "false");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_TTL", "120");
        Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_MAX_SIZE", "2000");

        try
        {
            var config = AuthConfig.LoadFromArgs([]);
            Assert.Equal(20, config.Jwt.ExternalTimeoutSeconds);
            Assert.False(config.Jwt.CacheEnabled);
            Assert.Equal(120, config.Jwt.CacheTtlSeconds);
            Assert.Equal(2000, config.Jwt.CacheMaxSize);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_TIMEOUT", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_ENABLED", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_TTL", null);
            Environment.SetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_MAX_SIZE", null);
        }
    }

    #endregion

    #region LoadFromCommandLine Additional Tests

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyEnabledAndDisabled()
    {
        var enabledConfig = AuthConfig.LoadFromArgs(["--auth-apikey-enabled"]);
        Assert.True(enabledConfig.ApiKey.Enabled);

        var disabledConfig = AuthConfig.LoadFromArgs(["--auth-apikey-disabled"]);
        Assert.False(disabledConfig.ApiKey.Enabled);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyModeWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-apikey-mode:introspection"]);
        Assert.Equal(ApiKeyMode.Introspection, colonConfig.ApiKey.Mode);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-apikey-mode=custom"]);
        Assert.Equal(ApiKeyMode.Custom, equalsConfig.ApiKey.Mode);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyKeysWithEquals()
    {
        var config = AuthConfig.LoadFromArgs(["--auth-apikey-keys=key1:tenant1,key2:tenant2"]);

        Assert.NotNull(config.ApiKey.Keys);
        Assert.Equal(2, config.ApiKey.Keys.Count);
        Assert.Equal("tenant1", config.ApiKey.Keys["key1"]);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyHeaderWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-apikey-header:X-Custom-Key"]);
        Assert.Equal("X-Custom-Key", colonConfig.ApiKey.HeaderName);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-apikey-header=X-Another-Key"]);
        Assert.Equal("X-Another-Key", equalsConfig.ApiKey.HeaderName);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyGroupHeaderWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-apikey-group-header:X-Custom-Group"]);
        Assert.Equal("X-Custom-Group", colonConfig.ApiKey.GroupIdentifierHeader);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-apikey-group-header=X-Another-Group"]);
        Assert.Equal("X-Another-Group", equalsConfig.ApiKey.GroupIdentifierHeader);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyIntrospectionAuthWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-apikey-introspection-auth:Bearer token1"]);
        Assert.Equal("Bearer token1", colonConfig.ApiKey.IntrospectionAuthHeader);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-apikey-introspection-auth=Bearer token2"]);
        Assert.Equal("Bearer token2", equalsConfig.ApiKey.IntrospectionAuthHeader);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyIntrospectionUrlWithEquals()
    {
        var config = AuthConfig.LoadFromArgs(["--auth-apikey-introspection-url=https://api.example.com/introspect"]);
        Assert.Equal("https://api.example.com/introspect", config.ApiKey.IntrospectionEndpoint);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyIntrospectionFieldWithEquals()
    {
        var config = AuthConfig.LoadFromArgs(["--auth-apikey-introspection-field=token"]);
        Assert.Equal("token", config.ApiKey.IntrospectionKeyField);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyTimeoutWithEquals()
    {
        var config = AuthConfig.LoadFromArgs(["--auth-apikey-timeout=25"]);
        Assert.Equal(25, config.ApiKey.ExternalTimeoutSeconds);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_ApiKeyCacheSettings()
    {
        var enabledConfig = AuthConfig.LoadFromArgs(["--auth-apikey-cache-enabled"]);
        Assert.True(enabledConfig.ApiKey.CacheEnabled);

        var disabledConfig = AuthConfig.LoadFromArgs(["--auth-apikey-cache-disabled"]);
        Assert.False(disabledConfig.ApiKey.CacheEnabled);

        var ttlColonConfig = AuthConfig.LoadFromArgs(["--auth-apikey-cache-ttl:500"]);
        Assert.Equal(500, ttlColonConfig.ApiKey.CacheTtlSeconds);

        var ttlEqualsConfig = AuthConfig.LoadFromArgs(["--auth-apikey-cache-ttl=600"]);
        Assert.Equal(600, ttlEqualsConfig.ApiKey.CacheTtlSeconds);

        var sizeColonConfig = AuthConfig.LoadFromArgs(["--auth-apikey-cache-max-size:8000"]);
        Assert.Equal(8000, sizeColonConfig.ApiKey.CacheMaxSize);

        var sizeEqualsConfig = AuthConfig.LoadFromArgs(["--auth-apikey-cache-max-size=9000"]);
        Assert.Equal(9000, sizeEqualsConfig.ApiKey.CacheMaxSize);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtEnabledAndDisabled()
    {
        var enabledConfig = AuthConfig.LoadFromArgs(["--auth-jwt-enabled"]);
        Assert.True(enabledConfig.Jwt.Enabled);

        var disabledConfig = AuthConfig.LoadFromArgs(["--auth-jwt-disabled"]);
        Assert.False(disabledConfig.Jwt.Enabled);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtModeWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-jwt-mode:gateway"]);
        Assert.Equal(JwtMode.Gateway, colonConfig.Jwt.Mode);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-jwt-mode=custom"]);
        Assert.Equal(JwtMode.Custom, equalsConfig.Jwt.Mode);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtSecretWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-jwt-secret:my-secret-1"]);
        Assert.Equal("my-secret-1", colonConfig.Jwt.Secret);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-jwt-secret=my-secret-2"]);
        Assert.Equal("my-secret-2", equalsConfig.Jwt.Secret);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtIssuerAndAudienceWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-jwt-issuer:issuer1", "--auth-jwt-audience:aud1"]);
        Assert.Equal("issuer1", colonConfig.Jwt.Issuer);
        Assert.Equal("aud1", colonConfig.Jwt.Audience);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-jwt-issuer=issuer2", "--auth-jwt-audience=aud2"]);
        Assert.Equal("issuer2", equalsConfig.Jwt.Issuer);
        Assert.Equal("aud2", equalsConfig.Jwt.Audience);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtClaimsWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-jwt-group-claim:org_id"]);
        Assert.Equal("org_id", colonConfig.Jwt.GroupIdentifierClaim);

        var equalsConfig = AuthConfig.LoadFromArgs(["--auth-jwt-group-claim=team_id"]);
        Assert.Equal("team_id", equalsConfig.Jwt.GroupIdentifierClaim);

        var userClaimConfig = AuthConfig.LoadFromArgs(["--auth-jwt-user-claim=uid"]);
        Assert.Equal("uid", userClaimConfig.Jwt.UserIdClaim);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtHeadersWithColonAndEquals()
    {
        var colonConfig = AuthConfig.LoadFromArgs(["--auth-jwt-group-header:X-Group", "--auth-jwt-user-header:X-User"]);
        Assert.Equal("X-Group", colonConfig.Jwt.GroupIdentifierHeader);
        Assert.Equal("X-User", colonConfig.Jwt.UserIdHeader);

        var equalsConfig =
            AuthConfig.LoadFromArgs(["--auth-jwt-group-header=X-Group2", "--auth-jwt-user-header=X-User2"]);
        Assert.Equal("X-Group2", equalsConfig.Jwt.GroupIdentifierHeader);
        Assert.Equal("X-User2", equalsConfig.Jwt.UserIdHeader);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtPublicKeyPathWithColon()
    {
        var config = AuthConfig.LoadFromArgs(["--auth-jwt-public-key-path:/etc/keys/public.pem"]);
        Assert.Equal("/etc/keys/public.pem", config.Jwt.PublicKeyPath);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtIntrospectionUrlWithEquals()
    {
        var config =
            AuthConfig.LoadFromArgs(["--auth-jwt-introspection-url=https://oauth.example.com/token/introspect"]);
        Assert.Equal("https://oauth.example.com/token/introspect", config.Jwt.IntrospectionEndpoint);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtClientIdAndSecretWithEquals()
    {
        var config =
            AuthConfig.LoadFromArgs(["--auth-jwt-client-id=client-abc", "--auth-jwt-client-secret=secret-xyz"]);
        Assert.Equal("client-abc", config.Jwt.ClientId);
        Assert.Equal("secret-xyz", config.Jwt.ClientSecret);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtTimeoutWithEquals()
    {
        var config = AuthConfig.LoadFromArgs(["--auth-jwt-timeout=30"]);
        Assert.Equal(30, config.Jwt.ExternalTimeoutSeconds);
    }

    [Fact]
    public void AuthConfig_LoadFromCommandLine_JwtCacheSettings()
    {
        var enabledConfig = AuthConfig.LoadFromArgs(["--auth-jwt-cache-enabled"]);
        Assert.True(enabledConfig.Jwt.CacheEnabled);

        var disabledConfig = AuthConfig.LoadFromArgs(["--auth-jwt-cache-disabled"]);
        Assert.False(disabledConfig.Jwt.CacheEnabled);

        var ttlColonConfig = AuthConfig.LoadFromArgs(["--auth-jwt-cache-ttl:400"]);
        Assert.Equal(400, ttlColonConfig.Jwt.CacheTtlSeconds);

        var ttlEqualsConfig = AuthConfig.LoadFromArgs(["--auth-jwt-cache-ttl=450"]);
        Assert.Equal(450, ttlEqualsConfig.Jwt.CacheTtlSeconds);

        var sizeColonConfig = AuthConfig.LoadFromArgs(["--auth-jwt-cache-max-size:7000"]);
        Assert.Equal(7000, sizeColonConfig.Jwt.CacheMaxSize);

        var sizeEqualsConfig = AuthConfig.LoadFromArgs(["--auth-jwt-cache-max-size=7500"]);
        Assert.Equal(7500, sizeEqualsConfig.Jwt.CacheMaxSize);
    }

    #endregion

    #region Validate Tests

    [Fact]
    public void AuthConfig_Validate_ApiKeyLocalModeWithoutKeys_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Local;
        config.ApiKey.Keys = null;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("Local mode requires at least one key", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ApiKeyLocalModeWithEmptyKeys_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Local;
        config.ApiKey.Keys = new Dictionary<string, string>();

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("Local mode requires at least one key", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ApiKeyIntrospectionModeWithoutEndpoint_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Introspection;
        config.ApiKey.IntrospectionEndpoint = null;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("Introspection mode requires", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ApiKeyCustomModeWithoutEndpoint_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Custom;
        config.ApiKey.CustomEndpoint = null;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("Custom mode requires", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ApiKeyInvalidTimeout_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Gateway;
        config.ApiKey.ExternalTimeoutSeconds = 0;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("timeout must be between 1 and 300", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ApiKeyTimeoutTooHigh_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Gateway;
        config.ApiKey.ExternalTimeoutSeconds = 301;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("timeout must be between 1 and 300", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ApiKeyCacheTtlInvalid_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Introspection;
        config.ApiKey.IntrospectionEndpoint = "https://example.com";
        config.ApiKey.CacheEnabled = true;
        config.ApiKey.CacheTtlSeconds = 0;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("cache TTL must be at least 1 second", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ApiKeyCacheMaxSizeInvalid_ShouldThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Custom;
        config.ApiKey.CustomEndpoint = "https://example.com";
        config.ApiKey.CacheEnabled = true;
        config.ApiKey.CacheTtlSeconds = 60;
        config.ApiKey.CacheMaxSize = 0;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("cache max size must be at least 1", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_JwtLocalModeWithoutSecretOrPublicKey_ShouldThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Local;
        config.Jwt.Secret = null;
        config.Jwt.PublicKeyPath = null;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("Local mode requires", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_JwtLocalModeWithNonExistentPublicKey_ShouldThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Local;
        config.Jwt.PublicKeyPath = "/nonexistent/path/to/key.pem";

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("public key file not found", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_JwtIntrospectionModeWithoutEndpoint_ShouldThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Introspection;
        config.Jwt.IntrospectionEndpoint = null;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("Introspection mode requires", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_JwtCustomModeWithoutEndpoint_ShouldThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Custom;
        config.Jwt.CustomEndpoint = null;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("Custom mode requires", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_JwtInvalidTimeout_ShouldThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Gateway;
        config.Jwt.ExternalTimeoutSeconds = -1;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("timeout must be between 1 and 300", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_JwtCacheTtlInvalid_ShouldThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Introspection;
        config.Jwt.IntrospectionEndpoint = "https://example.com";
        config.Jwt.CacheEnabled = true;
        config.Jwt.CacheTtlSeconds = 0;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("cache TTL must be at least 1 second", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_JwtCacheMaxSizeInvalid_ShouldThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Custom;
        config.Jwt.CustomEndpoint = "https://example.com";
        config.Jwt.CacheEnabled = true;
        config.Jwt.CacheTtlSeconds = 60;
        config.Jwt.CacheMaxSize = 0;

        var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());
        Assert.Contains("cache max size must be at least 1", ex.Message);
    }

    [Fact]
    public void AuthConfig_Validate_ValidApiKeyGatewayMode_ShouldNotThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Gateway;

        config.Validate();
    }

    [Fact]
    public void AuthConfig_Validate_ValidJwtGatewayMode_ShouldNotThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Gateway;

        config.Validate();
    }

    [Fact]
    public void AuthConfig_Validate_ValidApiKeyLocalModeWithKeys_ShouldNotThrow()
    {
        var config = new AuthConfig();
        config.ApiKey.Enabled = true;
        config.ApiKey.Mode = ApiKeyMode.Local;
        config.ApiKey.Keys = new Dictionary<string, string> { { "key1", "tenant1" } };

        config.Validate();
    }

    [Fact]
    public void AuthConfig_Validate_ValidJwtLocalModeWithSecret_ShouldNotThrow()
    {
        var config = new AuthConfig();
        config.Jwt.Enabled = true;
        config.Jwt.Mode = JwtMode.Local;
        config.Jwt.Secret = "my-secret-key";

        config.Validate();
    }

    #endregion
}
