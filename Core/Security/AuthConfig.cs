namespace AsposeMcpServer.Core.Security;

/// <summary>
///     API Key authentication verification mode
/// </summary>
public enum ApiKeyMode
{
    /// <summary>
    ///     Local verification using configured keys
    /// </summary>
    Local,

    /// <summary>
    ///     Trust API Gateway has verified the request
    /// </summary>
    Gateway,

    /// <summary>
    ///     Call external API to verify the key
    /// </summary>
    Introspection,

    /// <summary>
    ///     Call custom endpoint to verify
    /// </summary>
    Custom
}

/// <summary>
///     JWT authentication verification mode
/// </summary>
public enum JwtMode
{
    /// <summary>
    ///     Local verification using configured secret/public key
    /// </summary>
    Local,

    /// <summary>
    ///     Trust API Gateway has verified the token
    /// </summary>
    Gateway,

    /// <summary>
    ///     OAuth 2.0 Token Introspection
    /// </summary>
    Introspection,

    /// <summary>
    ///     Call custom endpoint to verify
    /// </summary>
    Custom
}

/// <summary>
///     API Key authentication configuration
/// </summary>
public class ApiKeyConfig
{
    /// <summary>
    ///     Enable API Key authentication
    /// </summary>
    public bool Enabled { get; set; }

    /// <summary>
    ///     Verification mode
    /// </summary>
    public ApiKeyMode Mode { get; set; } = ApiKeyMode.Local;

    /// <summary>
    ///     HTTP header name for API Key
    /// </summary>
    public string HeaderName { get; set; } = "X-API-Key";

    /// <summary>
    ///     Local mode: Dictionary of API keys to group identifiers
    /// </summary>
    public Dictionary<string, string>? Keys { get; set; }

    /// <summary>
    ///     Gateway mode: Header name containing group identifier (set by gateway)
    /// </summary>
    public string GroupIdentifierHeader { get; set; } = "X-Group-Id";

    /// <summary>
    ///     Introspection mode: Endpoint URL for key verification
    /// </summary>
    public string? IntrospectionEndpoint { get; set; }

    /// <summary>
    ///     Introspection mode: Authorization header value for introspection call
    /// </summary>
    public string? IntrospectionAuthHeader { get; set; }

    /// <summary>
    ///     Custom mode: Endpoint URL for custom verification
    ///     Should return JSON: { "valid": bool, "group_id": string }
    /// </summary>
    public string? CustomEndpoint { get; set; }

    /// <summary>
    ///     Timeout in seconds for external endpoint calls (Introspection/Custom mode)
    /// </summary>
    public int ExternalTimeoutSeconds { get; set; } = 5;

    /// <summary>
    ///     Introspection mode: Field name for API key in request body (default: "key")
    /// </summary>
    public string IntrospectionKeyField { get; set; } = "key";

    /// <summary>
    ///     Enable authentication result caching for Introspection/Custom modes.
    ///     When enabled, successful validation results are cached to reduce external API calls.
    /// </summary>
    public bool CacheEnabled { get; set; } = true;

    /// <summary>
    ///     Cache entry time-to-live in seconds (default: 300 = 5 minutes).
    ///     Shorter TTL improves security but increases external API calls.
    /// </summary>
    public int CacheTtlSeconds { get; set; } = 300;

    /// <summary>
    ///     Maximum number of cache entries (default: 10000).
    ///     Uses LRU eviction when limit is reached.
    /// </summary>
    public int CacheMaxSize { get; set; } = 10000;
}

/// <summary>
///     JWT authentication configuration
/// </summary>
public class JwtConfig
{
    /// <summary>
    ///     Enable JWT authentication
    /// </summary>
    public bool Enabled { get; set; }

    /// <summary>
    ///     Verification mode
    /// </summary>
    public JwtMode Mode { get; set; } = JwtMode.Local;

    /// <summary>
    ///     Local mode: Secret key for HMAC algorithms
    /// </summary>
    public string? Secret { get; set; }

    /// <summary>
    ///     Local mode: Path to public key file for RSA/ECDSA algorithms
    /// </summary>
    public string? PublicKeyPath { get; set; }

    /// <summary>
    ///     Local mode: Expected issuer claim
    /// </summary>
    public string? Issuer { get; set; }

    /// <summary>
    ///     Local mode: Expected audience claim
    /// </summary>
    public string? Audience { get; set; }

    /// <summary>
    ///     Claim name for group identifier (default: "tenant_id").
    ///     This can be configured to any claim such as "tenant_id", "team_id", "sub", or custom claims
    ///     depending on how you want to isolate sessions.
    /// </summary>
    public string GroupIdentifierClaim { get; set; } = "tenant_id";

    /// <summary>
    ///     Claim name for user ID, used for audit and logging (default: "sub")
    /// </summary>
    public string UserIdClaim { get; set; } = "sub";

    /// <summary>
    ///     Gateway mode: Header name containing group identifier (set by gateway)
    /// </summary>
    public string GroupIdentifierHeader { get; set; } = "X-Group-Id";

    /// <summary>
    ///     Gateway mode: Header name containing user ID (set by gateway)
    /// </summary>
    public string UserIdHeader { get; set; } = "X-User-Id";

    /// <summary>
    ///     Introspection mode: OAuth 2.0 Token Introspection endpoint
    /// </summary>
    public string? IntrospectionEndpoint { get; set; }

    /// <summary>
    ///     Introspection mode: OAuth client ID
    /// </summary>
    public string? ClientId { get; set; }

    /// <summary>
    ///     Introspection mode: OAuth client secret
    /// </summary>
    public string? ClientSecret { get; set; }

    /// <summary>
    ///     Custom mode: Endpoint URL for custom verification
    ///     Should return JSON: { "valid": bool, "group_id": string, "user_id": string }
    /// </summary>
    public string? CustomEndpoint { get; set; }

    /// <summary>
    ///     Timeout in seconds for external endpoint calls (Introspection/Custom mode)
    /// </summary>
    public int ExternalTimeoutSeconds { get; set; } = 5;

    /// <summary>
    ///     Enable authentication result caching for Introspection/Custom modes.
    ///     When enabled, successful validation results are cached to reduce external API calls.
    /// </summary>
    public bool CacheEnabled { get; set; } = true;

    /// <summary>
    ///     Cache entry time-to-live in seconds (default: 300 = 5 minutes).
    ///     Shorter TTL improves security but increases external API calls.
    /// </summary>
    public int CacheTtlSeconds { get; set; } = 300;

    /// <summary>
    ///     Maximum number of cache entries (default: 10000).
    ///     Uses LRU eviction when limit is reached.
    /// </summary>
    public int CacheMaxSize { get; set; } = 10000;
}

/// <summary>
///     Combined authentication configuration
/// </summary>
public class AuthConfig
{
    /// <summary>
    ///     API Key authentication settings
    /// </summary>
    public ApiKeyConfig ApiKey { get; set; } = new();

    /// <summary>
    ///     JWT authentication settings
    /// </summary>
    public JwtConfig Jwt { get; set; } = new();

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    ///     Command line arguments take precedence over environment variables.
    /// </summary>
    /// <param name="args">Command line arguments</param>
    /// <returns>AuthConfig instance</returns>
    public static AuthConfig LoadFromArgs(string[] args)
    {
        var config = new AuthConfig();
        config.LoadFromEnvironment();
        config.LoadFromCommandLine(args);
        return config;
    }

    /// <summary>
    ///     Loads configuration from environment variables
    /// </summary>
    private void LoadFromEnvironment()
    {
        LoadApiKeyFromEnvironment();
        LoadJwtFromEnvironment();
    }

    /// <summary>
    ///     Loads API Key configuration from environment variables
    /// </summary>
    private void LoadApiKeyFromEnvironment()
    {
        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_ENABLED"), out var apiKeyEnabled))
            ApiKey.Enabled = apiKeyEnabled;

        var apiKeyMode = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_MODE");
        if (!string.IsNullOrEmpty(apiKeyMode) && Enum.TryParse<ApiKeyMode>(apiKeyMode, true, out var parsedApiKeyMode))
            ApiKey.Mode = parsedApiKeyMode;

        var apiKeys = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_KEYS");
        if (!string.IsNullOrEmpty(apiKeys))
            ParseApiKeys(apiKeys);

        var apiKeyHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_HEADER");
        if (!string.IsNullOrEmpty(apiKeyHeader))
            ApiKey.HeaderName = apiKeyHeader;

        var apiKeyGroupHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_GROUP_HEADER");
        if (!string.IsNullOrEmpty(apiKeyGroupHeader))
            ApiKey.GroupIdentifierHeader = apiKeyGroupHeader;

        ApiKey.IntrospectionEndpoint =
            Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_URL");
        ApiKey.IntrospectionAuthHeader =
            Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_AUTH");
        ApiKey.CustomEndpoint = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CUSTOM_URL");

        var apiKeyIntrospectionField = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_INTROSPECTION_FIELD");
        if (!string.IsNullOrEmpty(apiKeyIntrospectionField))
            ApiKey.IntrospectionKeyField = apiKeyIntrospectionField;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_TIMEOUT"),
                out var apiKeyTimeout))
            ApiKey.ExternalTimeoutSeconds = apiKeyTimeout;

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_ENABLED"),
                out var apiKeyCacheEnabled))
            ApiKey.CacheEnabled = apiKeyCacheEnabled;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_TTL"),
                out var apiKeyCacheTtl))
            ApiKey.CacheTtlSeconds = apiKeyCacheTtl;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_CACHE_MAX_SIZE"),
                out var apiKeyCacheMaxSize))
            ApiKey.CacheMaxSize = apiKeyCacheMaxSize;
    }

    /// <summary>
    ///     Loads JWT configuration from environment variables
    /// </summary>
    private void LoadJwtFromEnvironment()
    {
        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_ENABLED"), out var jwtEnabled))
            Jwt.Enabled = jwtEnabled;

        var jwtMode = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_MODE");
        if (!string.IsNullOrEmpty(jwtMode) && Enum.TryParse<JwtMode>(jwtMode, true, out var parsedJwtMode))
            Jwt.Mode = parsedJwtMode;

        Jwt.Secret = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_SECRET");
        Jwt.PublicKeyPath = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_PUBLIC_KEY_PATH");
        Jwt.Issuer = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_ISSUER");
        Jwt.Audience = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_AUDIENCE");

        var jwtGroupClaim = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_GROUP_CLAIM");
        if (!string.IsNullOrEmpty(jwtGroupClaim))
            Jwt.GroupIdentifierClaim = jwtGroupClaim;

        var jwtUserClaim = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_CLAIM");
        if (!string.IsNullOrEmpty(jwtUserClaim))
            Jwt.UserIdClaim = jwtUserClaim;

        var jwtGroupHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_GROUP_HEADER");
        if (!string.IsNullOrEmpty(jwtGroupHeader))
            Jwt.GroupIdentifierHeader = jwtGroupHeader;

        var jwtUserHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_HEADER");
        if (!string.IsNullOrEmpty(jwtUserHeader))
            Jwt.UserIdHeader = jwtUserHeader;

        Jwt.IntrospectionEndpoint = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_INTROSPECTION_URL");
        Jwt.ClientId = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_ID");
        Jwt.ClientSecret = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_SECRET");
        Jwt.CustomEndpoint = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CUSTOM_URL");

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_TIMEOUT"), out var jwtTimeout))
            Jwt.ExternalTimeoutSeconds = jwtTimeout;

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_ENABLED"),
                out var jwtCacheEnabled))
            Jwt.CacheEnabled = jwtCacheEnabled;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_TTL"),
                out var jwtCacheTtl))
            Jwt.CacheTtlSeconds = jwtCacheTtl;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CACHE_MAX_SIZE"),
                out var jwtCacheMaxSize))
            Jwt.CacheMaxSize = jwtCacheMaxSize;
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables)
    /// </summary>
    /// <param name="args">Command line arguments</param>
    private void LoadFromCommandLine(string[] args)
    {
        foreach (var arg in args)
            if (!ProcessApiKeyArg(arg))
                _ = ProcessJwtArg(arg);
    }

    /// <summary>
    ///     Processes API Key related command line argument
    /// </summary>
    /// <returns>True if the argument was processed, false otherwise</returns>
    private bool ProcessApiKeyArg(string arg)
    {
        if (arg.Equals("--auth-apikey-enabled", StringComparison.OrdinalIgnoreCase))
        {
            ApiKey.Enabled = true;
            return true;
        }

        if (arg.Equals("--auth-apikey-disabled", StringComparison.OrdinalIgnoreCase))
        {
            ApiKey.Enabled = false;
            return true;
        }

        if (arg.Equals("--auth-apikey-cache-enabled", StringComparison.OrdinalIgnoreCase))
        {
            ApiKey.CacheEnabled = true;
            return true;
        }

        if (arg.Equals("--auth-apikey-cache-disabled", StringComparison.OrdinalIgnoreCase))
        {
            ApiKey.CacheEnabled = false;
            return true;
        }

        var value = TryGetArgValue(arg, "--auth-apikey-mode");
        if (value != null && Enum.TryParse<ApiKeyMode>(value, true, out var apiMode))
        {
            ApiKey.Mode = apiMode;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-keys");
        if (value != null)
        {
            ParseApiKeys(value);
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-header");
        if (value != null)
        {
            ApiKey.HeaderName = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-group-header");
        if (value != null)
        {
            ApiKey.GroupIdentifierHeader = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-introspection-auth");
        if (value != null)
        {
            ApiKey.IntrospectionAuthHeader = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-introspection-url");
        if (value != null)
        {
            ApiKey.IntrospectionEndpoint = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-custom-url");
        if (value != null)
        {
            ApiKey.CustomEndpoint = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-introspection-field");
        if (value != null)
        {
            ApiKey.IntrospectionKeyField = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-timeout");
        if (value != null && int.TryParse(value, out var timeout))
        {
            ApiKey.ExternalTimeoutSeconds = timeout;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-cache-ttl");
        if (value != null && int.TryParse(value, out var ttl))
        {
            ApiKey.CacheTtlSeconds = ttl;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-apikey-cache-max-size");
        if (value != null && int.TryParse(value, out var maxSize))
        {
            ApiKey.CacheMaxSize = maxSize;
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Processes JWT related command line argument
    /// </summary>
    /// <returns>True if the argument was processed, false otherwise</returns>
    private bool ProcessJwtArg(string arg)
    {
        if (arg.Equals("--auth-jwt-enabled", StringComparison.OrdinalIgnoreCase))
        {
            Jwt.Enabled = true;
            return true;
        }

        if (arg.Equals("--auth-jwt-disabled", StringComparison.OrdinalIgnoreCase))
        {
            Jwt.Enabled = false;
            return true;
        }

        if (arg.Equals("--auth-jwt-cache-enabled", StringComparison.OrdinalIgnoreCase))
        {
            Jwt.CacheEnabled = true;
            return true;
        }

        if (arg.Equals("--auth-jwt-cache-disabled", StringComparison.OrdinalIgnoreCase))
        {
            Jwt.CacheEnabled = false;
            return true;
        }

        var value = TryGetArgValue(arg, "--auth-jwt-mode");
        if (value != null && Enum.TryParse<JwtMode>(value, true, out var jwtMode))
        {
            Jwt.Mode = jwtMode;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-secret");
        if (value != null)
        {
            Jwt.Secret = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-issuer");
        if (value != null)
        {
            Jwt.Issuer = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-audience");
        if (value != null)
        {
            Jwt.Audience = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-group-claim");
        if (value != null)
        {
            Jwt.GroupIdentifierClaim = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-user-claim");
        if (value != null)
        {
            Jwt.UserIdClaim = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-group-header");
        if (value != null)
        {
            Jwt.GroupIdentifierHeader = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-user-header");
        if (value != null)
        {
            Jwt.UserIdHeader = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-public-key-path");
        if (value != null)
        {
            Jwt.PublicKeyPath = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-introspection-url");
        if (value != null)
        {
            Jwt.IntrospectionEndpoint = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-client-id");
        if (value != null)
        {
            Jwt.ClientId = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-client-secret");
        if (value != null)
        {
            Jwt.ClientSecret = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-custom-url");
        if (value != null)
        {
            Jwt.CustomEndpoint = value;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-timeout");
        if (value != null && int.TryParse(value, out var timeout))
        {
            Jwt.ExternalTimeoutSeconds = timeout;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-cache-ttl");
        if (value != null && int.TryParse(value, out var ttl))
        {
            Jwt.CacheTtlSeconds = ttl;
            return true;
        }

        value = TryGetArgValue(arg, "--auth-jwt-cache-max-size");
        if (value != null && int.TryParse(value, out var maxSize))
        {
            Jwt.CacheMaxSize = maxSize;
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Tries to extract value from command line argument with either ':' or '=' separator
    /// </summary>
    /// <param name="arg">The command line argument</param>
    /// <param name="prefix">The argument prefix (e.g., "--auth-jwt-secret")</param>
    /// <returns>The value if found, null otherwise</returns>
    private static string? TryGetArgValue(string arg, string prefix)
    {
        var colonPrefix = prefix + ":";
        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
            return arg[colonPrefix.Length..];

        var equalsPrefix = prefix + "=";
        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
            return arg[equalsPrefix.Length..];

        return null;
    }

    /// <summary>
    ///     Parses API key string in format "key1:groupId1,key2:groupId2"
    ///     Note: Only the first ':' is used as separator to allow ':' in API keys (e.g., base64 encoded keys)
    /// </summary>
    private void ParseApiKeys(string keysString)
    {
        ApiKey.Keys = new Dictionary<string, string>();
        foreach (var pair in keysString.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            // Split only on first ':' to allow ':' in API keys (e.g., base64 encoded keys)
            var separatorIndex = pair.IndexOf(':');
            if (separatorIndex > 0)
            {
                var key = pair[..separatorIndex].Trim();
                var groupId = pair[(separatorIndex + 1)..].Trim();
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(groupId))
                    ApiKey.Keys[key] = groupId;
            }
        }
    }

    /// <summary>
    ///     Validates the authentication configuration
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when configuration is invalid</exception>
    public void Validate()
    {
        if (ApiKey.Enabled)
        {
            switch (ApiKey.Mode)
            {
                case ApiKeyMode.Local when ApiKey.Keys is null or { Count: 0 }:
                    throw new InvalidOperationException(
                        "API Key authentication in Local mode requires at least one key configured via --auth-apikey-keys or ASPOSE_AUTH_APIKEY_KEYS");
                case ApiKeyMode.Introspection when string.IsNullOrEmpty(ApiKey.IntrospectionEndpoint):
                    throw new InvalidOperationException(
                        "API Key authentication in Introspection mode requires --auth-apikey-introspection-url or ASPOSE_AUTH_APIKEY_INTROSPECTION_URL");
                case ApiKeyMode.Custom when string.IsNullOrEmpty(ApiKey.CustomEndpoint):
                    throw new InvalidOperationException(
                        "API Key authentication in Custom mode requires --auth-apikey-custom-url or ASPOSE_AUTH_APIKEY_CUSTOM_URL");
            }

            ValidateExternalCallConfig("API Key", ApiKey.ExternalTimeoutSeconds, ApiKey.CacheEnabled,
                ApiKey.CacheTtlSeconds, ApiKey.CacheMaxSize,
                ApiKey.Mode == ApiKeyMode.Introspection || ApiKey.Mode == ApiKeyMode.Custom);
        }

        if (Jwt.Enabled)
        {
            switch (Jwt.Mode)
            {
                case JwtMode.Local when string.IsNullOrEmpty(Jwt.Secret) && string.IsNullOrEmpty(Jwt.PublicKeyPath):
                    throw new InvalidOperationException(
                        "JWT authentication in Local mode requires --auth-jwt-secret or --auth-jwt-public-key-path");
                case JwtMode.Local when !string.IsNullOrEmpty(Jwt.PublicKeyPath) && !File.Exists(Jwt.PublicKeyPath):
                    throw new InvalidOperationException(
                        $"JWT public key file not found: {Jwt.PublicKeyPath}");
                case JwtMode.Introspection when string.IsNullOrEmpty(Jwt.IntrospectionEndpoint):
                    throw new InvalidOperationException(
                        "JWT authentication in Introspection mode requires --auth-jwt-introspection-url or ASPOSE_AUTH_JWT_INTROSPECTION_URL");
                case JwtMode.Custom when string.IsNullOrEmpty(Jwt.CustomEndpoint):
                    throw new InvalidOperationException(
                        "JWT authentication in Custom mode requires --auth-jwt-custom-url or ASPOSE_AUTH_JWT_CUSTOM_URL");
            }

            ValidateExternalCallConfig("JWT", Jwt.ExternalTimeoutSeconds, Jwt.CacheEnabled,
                Jwt.CacheTtlSeconds, Jwt.CacheMaxSize,
                Jwt.Mode == JwtMode.Introspection || Jwt.Mode == JwtMode.Custom);
        }
    }

    /// <summary>
    ///     Validates external call configuration parameters
    /// </summary>
    // ReSharper disable once ParameterOnlyUsedForPreconditionCheck.Local - Validation is the intended purpose
    private static void ValidateExternalCallConfig(string authType, int timeoutSeconds, bool cacheEnabled,
        int cacheTtlSeconds, int cacheMaxSize, bool requiresExternalCall)
    {
        if (timeoutSeconds is < 1 or > 300)
            throw new InvalidOperationException(
                $"{authType} authentication timeout must be between 1 and 300 seconds");

        if (cacheEnabled && requiresExternalCall)
        {
            if (cacheTtlSeconds < 1)
                throw new InvalidOperationException(
                    $"{authType} cache TTL must be at least 1 second");
            if (cacheMaxSize < 1)
                throw new InvalidOperationException(
                    $"{authType} cache max size must be at least 1");
        }
    }
}
