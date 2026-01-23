namespace AsposeMcpServer.Core.Security;

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
