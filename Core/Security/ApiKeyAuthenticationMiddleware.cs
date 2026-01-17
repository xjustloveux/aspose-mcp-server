using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Authentication result from API Key validation
/// </summary>
public class ApiKeyAuthResult
{
    /// <summary>
    ///     Indicates whether the API key is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    ///     Group identifier extracted from the API key
    /// </summary>
    public string? GroupId { get; set; }

    /// <summary>
    ///     Error message if validation failed
    /// </summary>
    public string? ErrorMessage { get; set; }
}

/// <summary>
///     Middleware for API Key authentication supporting multiple verification modes
/// </summary>
public sealed class ApiKeyAuthenticationMiddleware : IMiddleware, IDisposable
{
    /// <summary>
    ///     Authentication result cache for Introspection/Custom modes
    /// </summary>
    private readonly AuthCache<ApiKeyAuthResult>? _cache;

    /// <summary>
    ///     API key authentication configuration
    /// </summary>
    private readonly ApiKeyConfig _config;

    /// <summary>
    ///     HTTP client for external validation calls
    /// </summary>
    private readonly HttpClient _httpClient;

    /// <summary>
    ///     Logger for authentication events
    /// </summary>
    private readonly ILogger<ApiKeyAuthenticationMiddleware> _logger;

    /// <summary>
    ///     Indicates whether we own the HTTP client and should dispose it
    /// </summary>
    private readonly bool _ownsHttpClient;

    /// <summary>
    ///     Tracks whether this middleware has been disposed (0 = not disposed, 1 = disposed)
    /// </summary>
    private int _disposed;

    /// <summary>
    ///     Creates a new API key authentication middleware instance
    /// </summary>
    /// <param name="config">API key configuration</param>
    /// <param name="logger">Logger instance</param>
    /// <param name="httpClientFactory">Optional HTTP client factory</param>
    public ApiKeyAuthenticationMiddleware(
        ApiKeyConfig config,
        ILogger<ApiKeyAuthenticationMiddleware> logger,
        IHttpClientFactory? httpClientFactory = null)
    {
        _config = config;
        _logger = logger;
        if (httpClientFactory != null)
        {
            _httpClient = httpClientFactory.CreateClient("ApiKeyAuth");
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

        if (config.CacheEnabled &&
            (config.Mode == ApiKeyMode.Introspection || config.Mode == ApiKeyMode.Custom))
            _cache = new AuthCache<ApiKeyAuthResult>(
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
        if (_ownsHttpClient)
            _httpClient.Dispose();
    }

    /// <summary>
    ///     Processes an HTTP request to validate API Key authentication
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

        var result = await ValidateApiKeyAsync(context);

        if (!result.IsValid)
        {
            _logger.LogWarning("API Key authentication failed: {Error}", result.ErrorMessage);
            context.Response.StatusCode = StatusCodes.Status401Unauthorized;
            context.Response.ContentType = "application/json";
            await context.Response.WriteAsync(JsonSerializer.Serialize(new
            {
                error = "Unauthorized",
                message = result.ErrorMessage ?? "Invalid or missing API key"
            }));
            return;
        }

        if (!string.IsNullOrEmpty(result.GroupId)) context.Items["GroupId"] = result.GroupId;

        await next(context);
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
    ///     Validates the API key from the request based on configured mode
    /// </summary>
    /// <param name="context">HTTP context containing the request</param>
    /// <returns>Authentication result with validation status</returns>
    private async Task<ApiKeyAuthResult> ValidateApiKeyAsync(HttpContext context)
    {
        var apiKey = context.Request.Headers[_config.HeaderName].FirstOrDefault();

        if (string.IsNullOrEmpty(apiKey))
            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = $"Missing {_config.HeaderName} header"
            };

        return _config.Mode switch
        {
            ApiKeyMode.Local => ValidateLocal(apiKey),
            ApiKeyMode.Gateway => ValidateGateway(context, apiKey),
            ApiKeyMode.Introspection => await ValidateWithCacheAsync(apiKey, () => ValidateIntrospectionAsync(apiKey)),
            ApiKeyMode.Custom => await ValidateWithCacheAsync(apiKey, () => ValidateCustomAsync(apiKey)),
            _ => new ApiKeyAuthResult { IsValid = false, ErrorMessage = "Unknown authentication mode" }
        };
    }

    /// <summary>
    ///     Validates API key with optional caching.
    ///     Only successful validation results are cached.
    /// </summary>
    /// <param name="apiKey">API key to use as cache key</param>
    /// <param name="validateFunc">Function to call when cache misses</param>
    /// <returns>Authentication result</returns>
    private async Task<ApiKeyAuthResult> ValidateWithCacheAsync(
        string apiKey,
        Func<Task<ApiKeyAuthResult>> validateFunc)
    {
        if (_cache == null)
            return await validateFunc();

        return await _cache.GetOrValidateAsync(
            apiKey,
            validateFunc,
            result => result.IsValid);
    }

    /// <summary>
    ///     Local mode: Validate API key against configured keys dictionary
    /// </summary>
    /// <param name="apiKey">API key to validate</param>
    /// <returns>Authentication result</returns>
    private ApiKeyAuthResult ValidateLocal(string apiKey)
    {
        if (_config.Keys == null || _config.Keys.Count == 0)
        {
            _logger.LogError("API Key authentication in Local mode but no keys configured");
            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = "Server configuration error: No API keys configured"
            };
        }

        if (_config.Keys.TryGetValue(apiKey, out var groupId))
        {
            _logger.LogDebug("API Key validated for group: {GroupId}", groupId);
            return new ApiKeyAuthResult
            {
                IsValid = true,
                GroupId = groupId
            };
        }

        return new ApiKeyAuthResult
        {
            IsValid = false,
            ErrorMessage = "Invalid API key"
        };
    }

    /// <summary>
    ///     Gateway mode: Trust that the API Gateway has validated the request
    ///     Extract group identifier from configured header
    /// </summary>
    /// <param name="context">HTTP context containing the request</param>
    /// <param name="_">API key (unused in gateway mode, validation is trusted from gateway)</param>
    /// <returns>Authentication result</returns>
    private ApiKeyAuthResult
        ValidateGateway(HttpContext context,
            string _) // NOSONAR S1172 - Discard pattern, parameter required by delegate signature
    {
        var groupId = context.Request.Headers[_config.GroupIdentifierHeader].FirstOrDefault();

        _logger.LogDebug("Gateway mode: Trusted request for group {GroupId}", groupId ?? "(anonymous)");
        return new ApiKeyAuthResult
        {
            IsValid = true,
            GroupId = groupId
        };
    }

    /// <summary>
    ///     Introspection mode: Call external API to validate the key
    /// </summary>
    /// <param name="apiKey">API key to validate</param>
    /// <returns>Authentication result</returns>
    private async Task<ApiKeyAuthResult> ValidateIntrospectionAsync(string apiKey)
    {
        if (string.IsNullOrEmpty(_config.IntrospectionEndpoint))
        {
            _logger.LogError("Introspection mode but no endpoint configured");
            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = "Server configuration error: No introspection endpoint configured"
            };
        }

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, _config.IntrospectionEndpoint);
            request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                [_config.IntrospectionKeyField] = apiKey
            });

            if (!string.IsNullOrEmpty(_config.IntrospectionAuthHeader))
                request.Headers.Authorization = AuthenticationHeaderValue.Parse(_config.IntrospectionAuthHeader);

            var response = await _httpClient.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Introspection endpoint returned {StatusCode}", response.StatusCode);
                return new ApiKeyAuthResult
                {
                    IsValid = false,
                    ErrorMessage = "API key validation failed"
                };
            }

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<IntrospectionResponse>(content, JsonDefaults.CaseInsensitive);

            if (result?.Active == true)
            {
                _logger.LogDebug("Introspection validated key for group: {GroupId}", result.GroupId);
                return new ApiKeyAuthResult
                {
                    IsValid = true,
                    GroupId = result.GroupId
                };
            }

            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = "API key is not active"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calling introspection endpoint");
            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = "Error validating API key"
            };
        }
    }

    /// <summary>
    ///     Custom mode: Call custom endpoint to validate the key
    /// </summary>
    /// <param name="apiKey">API key to validate</param>
    /// <returns>Authentication result</returns>
    private async Task<ApiKeyAuthResult> ValidateCustomAsync(string apiKey)
    {
        if (string.IsNullOrEmpty(_config.CustomEndpoint))
        {
            _logger.LogError("Custom mode but no endpoint configured");
            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = "Server configuration error: No custom endpoint configured"
            };
        }

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, _config.CustomEndpoint);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new { apiKey }),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Custom endpoint returned {StatusCode}", response.StatusCode);
                return new ApiKeyAuthResult
                {
                    IsValid = false,
                    ErrorMessage = "API key validation failed"
                };
            }

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<CustomValidationResponse>(content, JsonDefaults.CaseInsensitive);

            if (result?.Valid == true)
            {
                _logger.LogDebug("Custom endpoint validated key for group: {GroupId}", result.GroupId);
                return new ApiKeyAuthResult
                {
                    IsValid = true,
                    GroupId = result.GroupId
                };
            }

            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = result?.Error ?? "API key validation failed"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calling custom validation endpoint");
            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = "Error validating API key"
            };
        }
    }

    /// <summary>
    ///     Response from introspection endpoint
    /// </summary>
    private sealed class IntrospectionResponse
    {
        public bool Active { get; init; }

        [JsonPropertyName("group_id")] public string? GroupId { get; init; }
    }

    /// <summary>
    ///     Response from custom validation endpoint
    /// </summary>
    private sealed class CustomValidationResponse
    {
        public bool Valid { get; init; }

        [JsonPropertyName("group_id")] public string? GroupId { get; init; }

        public string? Error { get; init; }
    }
}
