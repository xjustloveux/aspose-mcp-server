namespace AsposeMcpServer.Core.Security;

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
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (!ProcessApiKeyArg(arg, args, ref i))
                ProcessJwtArg(arg, args, ref i);
        }
    }

    /// <summary>
    ///     Processes API Key related command line argument
    /// </summary>
    /// <param name="arg">The command line argument to process</param>
    /// <param name="args">All command line arguments</param>
    /// <param name="index">Current index in the arguments array</param>
    /// <returns>True if the argument was processed, false otherwise</returns>
    private bool ProcessApiKeyArg(string arg, string[] args, ref int index)
    {
        if (TryProcessApiKeyBooleanArg(arg)) return true;
        if (TryProcessApiKeyModeArg(arg, args, ref index)) return true;
        if (TryProcessApiKeyStringArg(arg, args, ref index)) return true;
        return TryProcessApiKeyIntArg(arg, args, ref index);
    }

    /// <summary>
    ///     Tries to process API key boolean arguments.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <returns>True if the argument was processed as a boolean API key argument; otherwise, false.</returns>
    private bool TryProcessApiKeyBooleanArg(string arg)
    {
        var booleanArgs = new Dictionary<string, Action>(StringComparer.OrdinalIgnoreCase)
        {
            ["--auth-apikey-enabled"] = () => ApiKey.Enabled = true,
            ["--auth-apikey-disabled"] = () => ApiKey.Enabled = false,
            ["--auth-apikey-cache-enabled"] = () => ApiKey.CacheEnabled = true,
            ["--auth-apikey-cache-disabled"] = () => ApiKey.CacheEnabled = false
        };

        if (!booleanArgs.TryGetValue(arg, out var action)) return false;
        action();
        return true;
    }

    /// <summary>
    ///     Tries to process API key mode argument.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <param name="args">All command line arguments.</param>
    /// <param name="index">Current index in the arguments array.</param>
    /// <returns>True if the argument was processed as an API key mode argument; otherwise, false.</returns>
    private bool TryProcessApiKeyModeArg(string arg, string[] args, ref int index)
    {
        var value = TryGetArgValue(arg, "--auth-apikey-mode", args, ref index);
        if (value == null || !Enum.TryParse<ApiKeyMode>(value, true, out var apiMode)) return false;
        ApiKey.Mode = apiMode;
        return true;
    }

    /// <summary>
    ///     Tries to process API key string arguments.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <param name="args">All command line arguments.</param>
    /// <param name="index">Current index in the arguments array.</param>
    /// <returns>True if the argument was processed as a string API key argument; otherwise, false.</returns>
    private bool TryProcessApiKeyStringArg(string arg, string[] args, ref int index)
    {
        var stringArgs = new Dictionary<string, Action<string>>
        {
            ["--auth-apikey-keys"] = ParseApiKeys,
            ["--auth-apikey-header"] = v => ApiKey.HeaderName = v,
            ["--auth-apikey-group-header"] = v => ApiKey.GroupIdentifierHeader = v,
            ["--auth-apikey-introspection-auth"] = v => ApiKey.IntrospectionAuthHeader = v,
            ["--auth-apikey-introspection-url"] = v => ApiKey.IntrospectionEndpoint = v,
            ["--auth-apikey-custom-url"] = v => ApiKey.CustomEndpoint = v,
            ["--auth-apikey-introspection-field"] = v => ApiKey.IntrospectionKeyField = v
        };

        foreach (var (prefix, action) in stringArgs)
        {
            var value = TryGetArgValue(arg, prefix, args, ref index);
            if (value == null) continue;
            action(value);
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Tries to process API key integer arguments.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <param name="args">All command line arguments.</param>
    /// <param name="index">Current index in the arguments array.</param>
    /// <returns>True if the argument was processed as an integer API key argument; otherwise, false.</returns>
    private bool TryProcessApiKeyIntArg(string arg, string[] args, ref int index)
    {
        var intArgs = new Dictionary<string, Action<int>>
        {
            ["--auth-apikey-timeout"] = v => ApiKey.ExternalTimeoutSeconds = v,
            ["--auth-apikey-cache-ttl"] = v => ApiKey.CacheTtlSeconds = v,
            ["--auth-apikey-cache-max-size"] = v => ApiKey.CacheMaxSize = v
        };

        foreach (var (prefix, action) in intArgs)
        {
            var value = TryGetArgValue(arg, prefix, args, ref index);
            if (value == null || !int.TryParse(value, out var intValue)) continue;
            action(intValue);
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Processes JWT related command line argument
    /// </summary>
    /// <param name="arg">The command line argument to process</param>
    /// <param name="args">All command line arguments</param>
    /// <param name="index">Current index in the arguments array</param>
    private void ProcessJwtArg(string arg, string[] args, ref int index)
    {
        if (TryProcessJwtBooleanArg(arg)) return;
        if (TryProcessJwtModeArg(arg, args, ref index)) return;
        if (TryProcessJwtStringArg(arg, args, ref index)) return;
        TryProcessJwtIntArg(arg, args, ref index);
    }

    /// <summary>
    ///     Tries to process JWT boolean arguments.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <returns>True if the argument was processed as a boolean JWT argument; otherwise, false.</returns>
    private bool TryProcessJwtBooleanArg(string arg)
    {
        var booleanArgs = new Dictionary<string, Action>(StringComparer.OrdinalIgnoreCase)
        {
            ["--auth-jwt-enabled"] = () => Jwt.Enabled = true,
            ["--auth-jwt-disabled"] = () => Jwt.Enabled = false,
            ["--auth-jwt-cache-enabled"] = () => Jwt.CacheEnabled = true,
            ["--auth-jwt-cache-disabled"] = () => Jwt.CacheEnabled = false
        };

        if (!booleanArgs.TryGetValue(arg, out var action)) return false;
        action();
        return true;
    }

    /// <summary>
    ///     Tries to process JWT mode argument.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <param name="args">All command line arguments.</param>
    /// <param name="index">Current index in the arguments array.</param>
    /// <returns>True if the argument was processed as a JWT mode argument; otherwise, false.</returns>
    private bool TryProcessJwtModeArg(string arg, string[] args, ref int index)
    {
        var value = TryGetArgValue(arg, "--auth-jwt-mode", args, ref index);
        if (value == null || !Enum.TryParse<JwtMode>(value, true, out var jwtMode)) return false;
        Jwt.Mode = jwtMode;
        return true;
    }

    /// <summary>
    ///     Tries to process JWT string arguments.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <param name="args">All command line arguments.</param>
    /// <param name="index">Current index in the arguments array.</param>
    /// <returns>True if the argument was processed as a string JWT argument; otherwise, false.</returns>
    private bool TryProcessJwtStringArg(string arg, string[] args, ref int index)
    {
        var stringArgs = new Dictionary<string, Action<string>>
        {
            ["--auth-jwt-secret"] = v => Jwt.Secret = v,
            ["--auth-jwt-issuer"] = v => Jwt.Issuer = v,
            ["--auth-jwt-audience"] = v => Jwt.Audience = v,
            ["--auth-jwt-group-claim"] = v => Jwt.GroupIdentifierClaim = v,
            ["--auth-jwt-user-claim"] = v => Jwt.UserIdClaim = v,
            ["--auth-jwt-group-header"] = v => Jwt.GroupIdentifierHeader = v,
            ["--auth-jwt-user-header"] = v => Jwt.UserIdHeader = v,
            ["--auth-jwt-public-key-path"] = v => Jwt.PublicKeyPath = v,
            ["--auth-jwt-introspection-url"] = v => Jwt.IntrospectionEndpoint = v,
            ["--auth-jwt-client-id"] = v => Jwt.ClientId = v,
            ["--auth-jwt-client-secret"] = v => Jwt.ClientSecret = v,
            ["--auth-jwt-custom-url"] = v => Jwt.CustomEndpoint = v
        };

        foreach (var (prefix, action) in stringArgs)
        {
            var value = TryGetArgValue(arg, prefix, args, ref index);
            if (value == null) continue;
            action(value);
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Tries to process JWT integer arguments.
    /// </summary>
    /// <param name="arg">The command line argument to process.</param>
    /// <param name="args">All command line arguments.</param>
    /// <param name="index">Current index in the arguments array.</param>
    private void TryProcessJwtIntArg(string arg, string[] args, ref int index)
    {
        var intArgs = new Dictionary<string, Action<int>>
        {
            ["--auth-jwt-timeout"] = v => Jwt.ExternalTimeoutSeconds = v,
            ["--auth-jwt-cache-ttl"] = v => Jwt.CacheTtlSeconds = v,
            ["--auth-jwt-cache-max-size"] = v => Jwt.CacheMaxSize = v
        };

        foreach (var (prefix, action) in intArgs)
        {
            var value = TryGetArgValue(arg, prefix, args, ref index);
            if (value == null || !int.TryParse(value, out var intValue)) continue;
            action(intValue);
            return;
        }
    }

    /// <summary>
    ///     Tries to extract value from command line argument with ':', '=' or space separator
    /// </summary>
    /// <param name="arg">The command line argument</param>
    /// <param name="prefix">The argument prefix (e.g., "--auth-jwt-secret")</param>
    /// <param name="args">All command line arguments</param>
    /// <param name="index">Current index in the arguments array</param>
    /// <returns>The value if found, null otherwise</returns>
    private static string? TryGetArgValue(string arg, string prefix, string[] args, ref int index)
    {
        if (arg.Equals(prefix, StringComparison.OrdinalIgnoreCase) && index + 1 < args.Length)
        {
            index++;
            return args[index];
        }

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
