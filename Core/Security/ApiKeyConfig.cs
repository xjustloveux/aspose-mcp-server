namespace AsposeMcpServer.Core.Security;

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
