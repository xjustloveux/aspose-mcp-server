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
    ///     Local mode: Dictionary of API keys to tenant IDs
    /// </summary>
    public Dictionary<string, string>? Keys { get; set; }

    /// <summary>
    ///     Gateway mode: Header name containing tenant ID (set by gateway)
    /// </summary>
    public string TenantIdHeader { get; set; } = "X-Tenant-Id";

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
    ///     Should return JSON: { "valid": bool, "tenant_id": string }
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
    ///     Local mode: Claim name for tenant ID (default: "tenant_id")
    /// </summary>
    public string TenantIdClaim { get; set; } = "tenant_id";

    /// <summary>
    ///     Local mode: Claim name for user ID (default: "sub")
    /// </summary>
    public string UserIdClaim { get; set; } = "sub";

    /// <summary>
    ///     Gateway mode: Header name containing tenant ID (set by gateway)
    /// </summary>
    public string TenantIdHeader { get; set; } = "X-Tenant-Id";

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
    ///     Should return JSON: { "valid": bool, "tenant_id": string, "user_id": string }
    /// </summary>
    public string? CustomEndpoint { get; set; }

    /// <summary>
    ///     Timeout in seconds for external endpoint calls (Introspection/Custom mode)
    /// </summary>
    public int ExternalTimeoutSeconds { get; set; } = 5;
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
        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_ENABLED"), out var apiKeyEnabled))
            ApiKey.Enabled = apiKeyEnabled;

        var apiKeyMode = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_MODE");
        if (!string.IsNullOrEmpty(apiKeyMode) && Enum.TryParse<ApiKeyMode>(apiKeyMode, true, out var parsedApiKeyMode))
            ApiKey.Mode = parsedApiKeyMode;

        var apiKeys = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_KEYS");
        if (!string.IsNullOrEmpty(apiKeys))
        {
            ApiKey.Keys = new Dictionary<string, string>();
            foreach (var pair in apiKeys.Split(',', StringSplitOptions.RemoveEmptyEntries))
            {
                // Split only on first ':' to allow ':' in API keys (e.g., base64 encoded keys)
                var separatorIndex = pair.IndexOf(':');
                if (separatorIndex > 0)
                {
                    var key = pair[..separatorIndex].Trim();
                    var tenant = pair[(separatorIndex + 1)..].Trim();
                    if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(tenant))
                        ApiKey.Keys[key] = tenant;
                }
            }
        }

        var apiKeyHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_HEADER");
        if (!string.IsNullOrEmpty(apiKeyHeader))
            ApiKey.HeaderName = apiKeyHeader;

        var apiKeyTenantHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_APIKEY_TENANT_HEADER");
        if (!string.IsNullOrEmpty(apiKeyTenantHeader))
            ApiKey.TenantIdHeader = apiKeyTenantHeader;

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

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_ENABLED"), out var jwtEnabled))
            Jwt.Enabled = jwtEnabled;

        var jwtMode = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_MODE");
        if (!string.IsNullOrEmpty(jwtMode) && Enum.TryParse<JwtMode>(jwtMode, true, out var parsedJwtMode))
            Jwt.Mode = parsedJwtMode;

        Jwt.Secret = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_SECRET");
        Jwt.PublicKeyPath = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_PUBLIC_KEY_PATH");
        Jwt.Issuer = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_ISSUER");
        Jwt.Audience = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_AUDIENCE");

        var jwtTenantClaim = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_TENANT_CLAIM");
        if (!string.IsNullOrEmpty(jwtTenantClaim))
            Jwt.TenantIdClaim = jwtTenantClaim;

        var jwtUserClaim = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_CLAIM");
        if (!string.IsNullOrEmpty(jwtUserClaim))
            Jwt.UserIdClaim = jwtUserClaim;

        var jwtTenantHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_TENANT_HEADER");
        if (!string.IsNullOrEmpty(jwtTenantHeader))
            Jwt.TenantIdHeader = jwtTenantHeader;

        var jwtUserHeader = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_USER_HEADER");
        if (!string.IsNullOrEmpty(jwtUserHeader))
            Jwt.UserIdHeader = jwtUserHeader;

        Jwt.IntrospectionEndpoint = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_INTROSPECTION_URL");
        Jwt.ClientId = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_ID");
        Jwt.ClientSecret = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CLIENT_SECRET");
        Jwt.CustomEndpoint = Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_CUSTOM_URL");

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_AUTH_JWT_TIMEOUT"), out var jwtTimeout))
            Jwt.ExternalTimeoutSeconds = jwtTimeout;
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables)
    /// </summary>
    /// <param name="args">Command line arguments</param>
    private void LoadFromCommandLine(string[] args)
    {
        foreach (var arg in args)
            if (arg.Equals("--auth-apikey-enabled", StringComparison.OrdinalIgnoreCase))
                ApiKey.Enabled = true;
            else if (arg.Equals("--auth-apikey-disabled", StringComparison.OrdinalIgnoreCase))
                ApiKey.Enabled = false;
            else if (arg.StartsWith("--auth-apikey-mode:", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<ApiKeyMode>(arg["--auth-apikey-mode:".Length..], true, out var apiMode1))
                ApiKey.Mode = apiMode1;
            else if (arg.StartsWith("--auth-apikey-mode=", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<ApiKeyMode>(arg["--auth-apikey-mode=".Length..], true, out var apiMode2))
                ApiKey.Mode = apiMode2;
            else if (arg.StartsWith("--auth-apikey-keys:", StringComparison.OrdinalIgnoreCase))
                ParseApiKeys(arg["--auth-apikey-keys:".Length..]);
            else if (arg.StartsWith("--auth-apikey-keys=", StringComparison.OrdinalIgnoreCase))
                ParseApiKeys(arg["--auth-apikey-keys=".Length..]);
            else if (arg.StartsWith("--auth-apikey-header:", StringComparison.OrdinalIgnoreCase))
                ApiKey.HeaderName = arg["--auth-apikey-header:".Length..];
            else if (arg.StartsWith("--auth-apikey-header=", StringComparison.OrdinalIgnoreCase))
                ApiKey.HeaderName = arg["--auth-apikey-header=".Length..];
            else if (arg.StartsWith("--auth-apikey-tenant-header:", StringComparison.OrdinalIgnoreCase))
                ApiKey.TenantIdHeader = arg["--auth-apikey-tenant-header:".Length..];
            else if (arg.StartsWith("--auth-apikey-tenant-header=", StringComparison.OrdinalIgnoreCase))
                ApiKey.TenantIdHeader = arg["--auth-apikey-tenant-header=".Length..];
            else if (arg.StartsWith("--auth-apikey-introspection-auth:", StringComparison.OrdinalIgnoreCase))
                ApiKey.IntrospectionAuthHeader = arg["--auth-apikey-introspection-auth:".Length..];
            else if (arg.StartsWith("--auth-apikey-introspection-auth=", StringComparison.OrdinalIgnoreCase))
                ApiKey.IntrospectionAuthHeader = arg["--auth-apikey-introspection-auth=".Length..];
            else if (arg.StartsWith("--auth-apikey-introspection-url:", StringComparison.OrdinalIgnoreCase))
                ApiKey.IntrospectionEndpoint = arg["--auth-apikey-introspection-url:".Length..];
            else if (arg.StartsWith("--auth-apikey-introspection-url=", StringComparison.OrdinalIgnoreCase))
                ApiKey.IntrospectionEndpoint = arg["--auth-apikey-introspection-url=".Length..];
            else if (arg.StartsWith("--auth-apikey-custom-url:", StringComparison.OrdinalIgnoreCase))
                ApiKey.CustomEndpoint = arg["--auth-apikey-custom-url:".Length..];
            else if (arg.StartsWith("--auth-apikey-custom-url=", StringComparison.OrdinalIgnoreCase))
                ApiKey.CustomEndpoint = arg["--auth-apikey-custom-url=".Length..];
            else if (arg.StartsWith("--auth-apikey-introspection-field:", StringComparison.OrdinalIgnoreCase))
                ApiKey.IntrospectionKeyField = arg["--auth-apikey-introspection-field:".Length..];
            else if (arg.StartsWith("--auth-apikey-introspection-field=", StringComparison.OrdinalIgnoreCase))
                ApiKey.IntrospectionKeyField = arg["--auth-apikey-introspection-field=".Length..];
            else if (arg.StartsWith("--auth-apikey-timeout:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--auth-apikey-timeout:".Length..], out var apiKeyTimeout1))
                ApiKey.ExternalTimeoutSeconds = apiKeyTimeout1;
            else if (arg.StartsWith("--auth-apikey-timeout=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--auth-apikey-timeout=".Length..], out var apiKeyTimeout2))
                ApiKey.ExternalTimeoutSeconds = apiKeyTimeout2;
            else if (arg.Equals("--auth-jwt-enabled", StringComparison.OrdinalIgnoreCase))
                Jwt.Enabled = true;
            else if (arg.Equals("--auth-jwt-disabled", StringComparison.OrdinalIgnoreCase))
                Jwt.Enabled = false;
            else if (arg.StartsWith("--auth-jwt-mode:", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<JwtMode>(arg["--auth-jwt-mode:".Length..], true, out var jwtMode1))
                Jwt.Mode = jwtMode1;
            else if (arg.StartsWith("--auth-jwt-mode=", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<JwtMode>(arg["--auth-jwt-mode=".Length..], true, out var jwtMode2))
                Jwt.Mode = jwtMode2;
            else if (arg.StartsWith("--auth-jwt-secret:", StringComparison.OrdinalIgnoreCase))
                Jwt.Secret = arg["--auth-jwt-secret:".Length..];
            else if (arg.StartsWith("--auth-jwt-secret=", StringComparison.OrdinalIgnoreCase))
                Jwt.Secret = arg["--auth-jwt-secret=".Length..];
            else if (arg.StartsWith("--auth-jwt-issuer:", StringComparison.OrdinalIgnoreCase))
                Jwt.Issuer = arg["--auth-jwt-issuer:".Length..];
            else if (arg.StartsWith("--auth-jwt-issuer=", StringComparison.OrdinalIgnoreCase))
                Jwt.Issuer = arg["--auth-jwt-issuer=".Length..];
            else if (arg.StartsWith("--auth-jwt-audience:", StringComparison.OrdinalIgnoreCase))
                Jwt.Audience = arg["--auth-jwt-audience:".Length..];
            else if (arg.StartsWith("--auth-jwt-audience=", StringComparison.OrdinalIgnoreCase))
                Jwt.Audience = arg["--auth-jwt-audience=".Length..];
            else if (arg.StartsWith("--auth-jwt-tenant-claim:", StringComparison.OrdinalIgnoreCase))
                Jwt.TenantIdClaim = arg["--auth-jwt-tenant-claim:".Length..];
            else if (arg.StartsWith("--auth-jwt-tenant-claim=", StringComparison.OrdinalIgnoreCase))
                Jwt.TenantIdClaim = arg["--auth-jwt-tenant-claim=".Length..];
            else if (arg.StartsWith("--auth-jwt-user-claim:", StringComparison.OrdinalIgnoreCase))
                Jwt.UserIdClaim = arg["--auth-jwt-user-claim:".Length..];
            else if (arg.StartsWith("--auth-jwt-user-claim=", StringComparison.OrdinalIgnoreCase))
                Jwt.UserIdClaim = arg["--auth-jwt-user-claim=".Length..];
            else if (arg.StartsWith("--auth-jwt-tenant-header:", StringComparison.OrdinalIgnoreCase))
                Jwt.TenantIdHeader = arg["--auth-jwt-tenant-header:".Length..];
            else if (arg.StartsWith("--auth-jwt-tenant-header=", StringComparison.OrdinalIgnoreCase))
                Jwt.TenantIdHeader = arg["--auth-jwt-tenant-header=".Length..];
            else if (arg.StartsWith("--auth-jwt-user-header:", StringComparison.OrdinalIgnoreCase))
                Jwt.UserIdHeader = arg["--auth-jwt-user-header:".Length..];
            else if (arg.StartsWith("--auth-jwt-user-header=", StringComparison.OrdinalIgnoreCase))
                Jwt.UserIdHeader = arg["--auth-jwt-user-header=".Length..];
            else if (arg.StartsWith("--auth-jwt-public-key-path:", StringComparison.OrdinalIgnoreCase))
                Jwt.PublicKeyPath = arg["--auth-jwt-public-key-path:".Length..];
            else if (arg.StartsWith("--auth-jwt-public-key-path=", StringComparison.OrdinalIgnoreCase))
                Jwt.PublicKeyPath = arg["--auth-jwt-public-key-path=".Length..];
            else if (arg.StartsWith("--auth-jwt-introspection-url:", StringComparison.OrdinalIgnoreCase))
                Jwt.IntrospectionEndpoint = arg["--auth-jwt-introspection-url:".Length..];
            else if (arg.StartsWith("--auth-jwt-introspection-url=", StringComparison.OrdinalIgnoreCase))
                Jwt.IntrospectionEndpoint = arg["--auth-jwt-introspection-url=".Length..];
            else if (arg.StartsWith("--auth-jwt-client-id:", StringComparison.OrdinalIgnoreCase))
                Jwt.ClientId = arg["--auth-jwt-client-id:".Length..];
            else if (arg.StartsWith("--auth-jwt-client-id=", StringComparison.OrdinalIgnoreCase))
                Jwt.ClientId = arg["--auth-jwt-client-id=".Length..];
            else if (arg.StartsWith("--auth-jwt-client-secret:", StringComparison.OrdinalIgnoreCase))
                Jwt.ClientSecret = arg["--auth-jwt-client-secret:".Length..];
            else if (arg.StartsWith("--auth-jwt-client-secret=", StringComparison.OrdinalIgnoreCase))
                Jwt.ClientSecret = arg["--auth-jwt-client-secret=".Length..];
            else if (arg.StartsWith("--auth-jwt-custom-url:", StringComparison.OrdinalIgnoreCase))
                Jwt.CustomEndpoint = arg["--auth-jwt-custom-url:".Length..];
            else if (arg.StartsWith("--auth-jwt-custom-url=", StringComparison.OrdinalIgnoreCase))
                Jwt.CustomEndpoint = arg["--auth-jwt-custom-url=".Length..];
            else if (arg.StartsWith("--auth-jwt-timeout:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--auth-jwt-timeout:".Length..], out var jwtTimeout1))
                Jwt.ExternalTimeoutSeconds = jwtTimeout1;
            else if (arg.StartsWith("--auth-jwt-timeout=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--auth-jwt-timeout=".Length..], out var jwtTimeout2))
                Jwt.ExternalTimeoutSeconds = jwtTimeout2;
    }

    /// <summary>
    ///     Parses API key string in format "key1:tenant1,key2:tenant2"
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
                var tenant = pair[(separatorIndex + 1)..].Trim();
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(tenant))
                    ApiKey.Keys[key] = tenant;
            }
        }
    }
}