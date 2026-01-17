using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using AsposeMcpServer.Core.Helpers;
using Microsoft.IdentityModel.Tokens;

namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Authentication result from JWT validation
/// </summary>
public class JwtAuthResult
{
    /// <summary>
    ///     Indicates whether the JWT token is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    ///     Group identifier extracted from the token
    /// </summary>
    public string? GroupId { get; set; }

    /// <summary>
    ///     User identifier extracted from the token (for audit and logging)
    /// </summary>
    public string? UserId { get; set; }

    /// <summary>
    ///     Error message if validation failed
    /// </summary>
    public string? ErrorMessage { get; set; }
}

/// <summary>
///     Middleware for JWT authentication supporting multiple verification modes
/// </summary>
public sealed class JwtAuthenticationMiddleware : IMiddleware, IDisposable
{
    /// <summary>
    ///     Authentication result cache for Introspection/Custom modes
    /// </summary>
    private readonly AuthCache<JwtAuthResult>? _cache;

    /// <summary>
    ///     JWT authentication configuration
    /// </summary>
    private readonly JwtConfig _config;

    /// <summary>
    ///     Cryptographic key instance (RSA or ECDsa) that needs to be disposed
    /// </summary>
    private readonly IDisposable? _cryptoKey;

    /// <summary>
    ///     HTTP client for external validation calls
    /// </summary>
    private readonly HttpClient _httpClient;

    /// <summary>
    ///     Logger for authentication events
    /// </summary>
    private readonly ILogger<JwtAuthenticationMiddleware> _logger;

    /// <summary>
    ///     Indicates whether we own the HTTP client and should dispose it
    /// </summary>
    private readonly bool _ownsHttpClient;

    /// <summary>
    ///     Token validation parameters for local mode
    /// </summary>
    private readonly TokenValidationParameters? _validationParameters;

    /// <summary>
    ///     Tracks whether this middleware has been disposed (0 = not disposed, 1 = disposed)
    /// </summary>
    private int _disposed;

    /// <summary>
    ///     Creates a new JWT authentication middleware instance
    /// </summary>
    /// <param name="config">JWT configuration</param>
    /// <param name="logger">Logger instance</param>
    /// <param name="httpClientFactory">Optional HTTP client factory</param>
    public JwtAuthenticationMiddleware(
        JwtConfig config,
        ILogger<JwtAuthenticationMiddleware> logger,
        IHttpClientFactory? httpClientFactory = null)
    {
        _config = config;
        _logger = logger;
        if (httpClientFactory != null)
        {
            _httpClient = httpClientFactory.CreateClient("JwtAuth");
            _httpClient.Timeout = TimeSpan.FromSeconds(config.ExternalTimeoutSeconds);
            _ownsHttpClient = false;
        }
        else
        {
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(config.ExternalTimeoutSeconds)
            };
            _ownsHttpClient = true;
        }

        if (config.Mode == JwtMode.Local)
            (_validationParameters, _cryptoKey) = BuildValidationParameters();

        if (config.CacheEnabled &&
            (config.Mode == JwtMode.Introspection || config.Mode == JwtMode.Custom))
            _cache = new AuthCache<JwtAuthResult>(
                config.CacheTtlSeconds,
                config.CacheMaxSize);
    }

    /// <summary>
    ///     Releases resources used by this middleware.
    ///     Thread-safe: uses Interlocked to prevent double-dispose.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
            return;

        _cache?.Clear();
        _cryptoKey?.Dispose();
        if (_ownsHttpClient)
            _httpClient.Dispose();
    }

    /// <summary>
    ///     Processes an HTTP request to validate JWT authentication
    /// </summary>
    /// <param name="context">HTTP context for the current request</param>
    /// <param name="next">Next middleware delegate</param>
    public async Task InvokeAsync(HttpContext context, RequestDelegate next)
    {
        if (ShouldSkipAuthentication(context.Request.Path))
        {
            await next(context);
            return;
        }

        var result = await ValidateJwtAsync(context);

        if (!result.IsValid)
        {
            _logger.LogWarning("JWT authentication failed: {Error}", result.ErrorMessage);
            context.Response.StatusCode = StatusCodes.Status401Unauthorized;
            context.Response.ContentType = "application/json";
            await context.Response.WriteAsync(JsonSerializer.Serialize(new
            {
                error = "Unauthorized",
                message = result.ErrorMessage ?? "Invalid or missing JWT token"
            }));
            return;
        }

        if (!string.IsNullOrEmpty(result.GroupId))
            context.Items["GroupId"] = result.GroupId;
        if (!string.IsNullOrEmpty(result.UserId))
            context.Items["UserId"] = result.UserId;

        await next(context);
    }

    /// <summary>
    ///     Builds token validation parameters from configuration
    /// </summary>
    /// <returns>Token validation parameters and disposable crypto key (if any)</returns>
    private (TokenValidationParameters?, IDisposable?) BuildValidationParameters()
    {
        SecurityKey? signingKey = null;
        IDisposable? cryptoKey = null;

        if (!string.IsNullOrEmpty(_config.Secret))
        {
            signingKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_config.Secret));
        }
        else if (!string.IsNullOrEmpty(_config.PublicKeyPath) && File.Exists(_config.PublicKeyPath))
        {
            var keyText = File.ReadAllText(_config.PublicKeyPath);
            if (keyText.Contains("RSA PUBLIC KEY") || keyText.Contains("PUBLIC KEY"))
            {
                var rsa = RSA.Create();
                rsa.ImportFromPem(keyText);
                signingKey = new RsaSecurityKey(rsa);
                cryptoKey = rsa;
            }
            else if (keyText.Contains("EC PUBLIC KEY"))
            {
                var ecdsa = ECDsa.Create();
                ecdsa.ImportFromPem(keyText);
                signingKey = new ECDsaSecurityKey(ecdsa);
                cryptoKey = ecdsa;
            }
        }

        if (signingKey == null)
        {
            _logger.LogWarning("JWT Local mode: No valid signing key configured");
            return (null, null);
        }

        return (new TokenValidationParameters
        {
            ValidateIssuerSigningKey = true,
            IssuerSigningKey = signingKey,
            ValidateIssuer = !string.IsNullOrEmpty(_config.Issuer),
            ValidIssuer = _config.Issuer,
            ValidateAudience = !string.IsNullOrEmpty(_config.Audience),
            ValidAudience = _config.Audience,
            ValidateLifetime = true,
            ClockSkew = TimeSpan.FromMinutes(5)
        }, cryptoKey);
    }

    /// <summary>
    ///     Determines if authentication should be skipped for health/metrics endpoints
    /// </summary>
    /// <param name="path">Request path to check</param>
    /// <returns>True if authentication should be skipped</returns>
    private static bool ShouldSkipAuthentication(PathString path)
    {
        return path.StartsWithSegments("/health") ||
               path.StartsWithSegments("/metrics") ||
               path.StartsWithSegments("/ready");
    }

    /// <summary>
    ///     Validates the JWT token from the request based on configured mode
    /// </summary>
    /// <param name="context">HTTP context containing the request</param>
    /// <returns>Authentication result with validation status</returns>
    private async Task<JwtAuthResult> ValidateJwtAsync(HttpContext context)
    {
        var authHeader = context.Request.Headers.Authorization.FirstOrDefault();
        if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Missing or invalid Authorization header"
            };

        var token = authHeader["Bearer ".Length..].Trim();

        return _config.Mode switch
        {
            JwtMode.Local => ValidateLocal(token),
            JwtMode.Gateway => ValidateGateway(context),
            JwtMode.Introspection => await ValidateWithCacheAsync(token, () => ValidateIntrospectionAsync(token)),
            JwtMode.Custom => await ValidateWithCacheAsync(token, () => ValidateCustomAsync(token)),
            _ => new JwtAuthResult { IsValid = false, ErrorMessage = "Unknown authentication mode" }
        };
    }

    /// <summary>
    ///     Validates JWT token with optional caching.
    ///     Only successful validation results are cached.
    /// </summary>
    /// <param name="token">Token to use as cache key</param>
    /// <param name="validateFunc">Function to call when cache misses</param>
    /// <returns>Authentication result</returns>
    private async Task<JwtAuthResult> ValidateWithCacheAsync(
        string token,
        Func<Task<JwtAuthResult>> validateFunc)
    {
        if (_cache == null)
            return await validateFunc();

        return await _cache.GetOrValidateAsync(
            token,
            validateFunc,
            result => result.IsValid);
    }

    /// <summary>
    ///     Local mode: Validate JWT using configured secret/public key
    /// </summary>
    /// <param name="token">JWT token to validate</param>
    /// <returns>Authentication result</returns>
    private JwtAuthResult ValidateLocal(string token)
    {
        if (_validationParameters == null)
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Server configuration error: No JWT signing key configured"
            };

        try
        {
            var handler = new JwtSecurityTokenHandler
            {
                MapInboundClaims = false
            };
            var principal = handler.ValidateToken(token, _validationParameters, out _);

            var groupId = principal.FindFirst(_config.GroupIdentifierClaim)?.Value;
            var userId = principal.FindFirst(_config.UserIdClaim)?.Value;

            _logger.LogDebug("JWT validated for group: {GroupId}, user: {UserId}", groupId, userId);

            return new JwtAuthResult
            {
                IsValid = true,
                GroupId = groupId,
                UserId = userId
            };
        }
        catch (SecurityTokenExpiredException)
        {
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Token has expired"
            };
        }
        catch (SecurityTokenException ex)
        {
            _logger.LogDebug("Token validation failed: {Message}",
                ex.Message); // NOSONAR S6667 - Structured logging with placeholders is correct pattern
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Invalid token"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error validating JWT");
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Error validating token"
            };
        }
    }

    /// <summary>
    ///     Gateway mode: Trust that the API Gateway has validated the token
    ///     Extract group/user ID from configured headers
    /// </summary>
    /// <param name="context">HTTP context containing the request</param>
    /// <returns>Authentication result</returns>
    private JwtAuthResult ValidateGateway(HttpContext context)
    {
        var groupId = context.Request.Headers[_config.GroupIdentifierHeader].FirstOrDefault();
        var userId = context.Request.Headers[_config.UserIdHeader].FirstOrDefault();

        _logger.LogDebug("Gateway mode: Trusted request for group {GroupId}, user {UserId}",
            groupId ?? "(anonymous)", userId ?? "(anonymous)");

        return new JwtAuthResult
        {
            IsValid = true,
            GroupId = groupId,
            UserId = userId
        };
    }

    /// <summary>
    ///     Introspection mode: OAuth 2.0 Token Introspection (RFC 7662)
    /// </summary>
    /// <param name="token">JWT token to validate</param>
    /// <returns>Authentication result</returns>
    private async Task<JwtAuthResult> ValidateIntrospectionAsync(string token)
    {
        if (string.IsNullOrEmpty(_config.IntrospectionEndpoint))
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Server configuration error: No introspection endpoint configured"
            };

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, _config.IntrospectionEndpoint);

            if (!string.IsNullOrEmpty(_config.ClientId) && !string.IsNullOrEmpty(_config.ClientSecret))
            {
                var credentials = Convert.ToBase64String(
                    Encoding.UTF8.GetBytes($"{_config.ClientId}:{_config.ClientSecret}"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            }

            request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["token"] = token,
                ["token_type_hint"] = "access_token"
            });

            var response = await _httpClient.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Introspection endpoint returned {StatusCode}", response.StatusCode);
                return new JwtAuthResult
                {
                    IsValid = false,
                    ErrorMessage = "Token validation failed"
                };
            }

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<IntrospectionResponse>(content, JsonDefaults.CaseInsensitive);

            if (result?.Active == true)
            {
                _logger.LogDebug("Introspection validated token for sub: {Sub}", result.Sub);
                return new JwtAuthResult
                {
                    IsValid = true,
                    GroupId = result.GroupId ?? result.ClientId,
                    UserId = result.Sub
                };
            }

            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Token is not active"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calling introspection endpoint");
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Error validating token"
            };
        }
    }

    /// <summary>
    ///     Custom mode: Call custom endpoint to validate the token
    /// </summary>
    /// <param name="token">JWT token to validate</param>
    /// <returns>Authentication result</returns>
    private async Task<JwtAuthResult> ValidateCustomAsync(string token)
    {
        if (string.IsNullOrEmpty(_config.CustomEndpoint))
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Server configuration error: No custom endpoint configured"
            };

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, _config.CustomEndpoint);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new { token }),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Custom endpoint returned {StatusCode}", response.StatusCode);
                return new JwtAuthResult
                {
                    IsValid = false,
                    ErrorMessage = "Token validation failed"
                };
            }

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<CustomValidationResponse>(content, JsonDefaults.CaseInsensitive);

            if (result?.Valid == true)
            {
                _logger.LogDebug("Custom endpoint validated token for group: {GroupId}", result.GroupId);
                return new JwtAuthResult
                {
                    IsValid = true,
                    GroupId = result.GroupId,
                    UserId = result.UserId
                };
            }

            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = result?.Error ?? "Token validation failed"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calling custom validation endpoint");
            return new JwtAuthResult
            {
                IsValid = false,
                ErrorMessage = "Error validating token"
            };
        }
    }

    /// <summary>
    ///     Response from OAuth 2.0 introspection endpoint (RFC 7662)
    /// </summary>
    private sealed class IntrospectionResponse // NOSONAR S3459 - Properties assigned via JSON deserialization
    {
        public bool Active { get; init; }
        public string? Sub { get; init; }

        [JsonPropertyName("client_id")] public string? ClientId { get; init; }

        [JsonPropertyName("group_id")] public string? GroupId { get; init; }
    }

    /// <summary>
    ///     Response from custom validation endpoint
    /// </summary>
    private sealed class CustomValidationResponse // NOSONAR S3459 - Properties assigned via JSON deserialization
    {
        public bool Valid { get; init; }

        [JsonPropertyName("group_id")] public string? GroupId { get; init; }

        [JsonPropertyName("user_id")] public string? UserId { get; init; }

        public string? Error { get; init; }
    }
}
