using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

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
    ///     Tenant identifier extracted from the API key
    /// </summary>
    public string? TenantId { get; set; }

    /// <summary>
    ///     Error message if validation failed
    /// </summary>
    public string? ErrorMessage { get; set; }
}

/// <summary>
///     Middleware for API Key authentication supporting multiple verification modes
/// </summary>
public class ApiKeyAuthenticationMiddleware : IDisposable
{
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
    ///     Next middleware in the pipeline
    /// </summary>
    private readonly RequestDelegate _next;

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
    /// <param name="next">Next middleware delegate</param>
    /// <param name="config">API key configuration</param>
    /// <param name="logger">Logger instance</param>
    /// <param name="httpClientFactory">Optional HTTP client factory</param>
    public ApiKeyAuthenticationMiddleware(
        RequestDelegate next,
        ApiKeyConfig config,
        ILogger<ApiKeyAuthenticationMiddleware> logger,
        IHttpClientFactory? httpClientFactory = null)
    {
        _next = next;
        _config = config;
        _logger = logger;
        if (httpClientFactory != null)
        {
            _httpClient = httpClientFactory.CreateClient("ApiKeyAuth");
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
    }

    /// <summary>
    ///     Releases resources used by this middleware.
    ///     Thread-safe: uses Interlocked to prevent double-dispose.
    /// </summary>
    public void Dispose()
    {
        // Atomically set _disposed to 1, return previous value
        // If previous value was already 1, another thread already disposed
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
            return;

        if (_ownsHttpClient)
            _httpClient.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    ///     Processes an HTTP request to validate API Key authentication
    /// </summary>
    /// <param name="context">HTTP context for the current request</param>
    public async Task InvokeAsync(HttpContext context)
    {
        if (ShouldSkipAuthentication(context.Request.Path))
        {
            await _next(context);
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

        if (!string.IsNullOrEmpty(result.TenantId)) context.Items["TenantId"] = result.TenantId;

        await _next(context);
    }

    /// <summary>
    ///     Determines if authentication should be skipped for health/metrics endpoints
    /// </summary>
    /// <param name="path">Request path to check</param>
    /// <returns>True if authentication should be skipped</returns>
    private bool ShouldSkipAuthentication(PathString path)
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
            ApiKeyMode.Introspection => await ValidateIntrospectionAsync(apiKey),
            ApiKeyMode.Custom => await ValidateCustomAsync(apiKey),
            _ => new ApiKeyAuthResult { IsValid = false, ErrorMessage = "Unknown authentication mode" }
        };
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

        if (_config.Keys.TryGetValue(apiKey, out var tenantId))
        {
            _logger.LogDebug("API Key validated for tenant: {TenantId}", tenantId);
            return new ApiKeyAuthResult
            {
                IsValid = true,
                TenantId = tenantId
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
    ///     Extract tenant ID from configured header
    /// </summary>
    /// <param name="context">HTTP context containing the request</param>
    /// <param name="_">API key (unused in gateway mode, validation is trusted from gateway)</param>
    /// <returns>Authentication result</returns>
    private ApiKeyAuthResult ValidateGateway(HttpContext context, string _)
    {
        var tenantId = context.Request.Headers[_config.TenantIdHeader].FirstOrDefault();

        if (string.IsNullOrEmpty(tenantId))
        {
            _logger.LogWarning("Gateway mode: Missing tenant ID header {Header}", _config.TenantIdHeader);
            return new ApiKeyAuthResult
            {
                IsValid = false,
                ErrorMessage = $"Missing {_config.TenantIdHeader} header from gateway"
            };
        }

        _logger.LogDebug("Gateway mode: Trusted request for tenant {TenantId}", tenantId);
        return new ApiKeyAuthResult
        {
            IsValid = true,
            TenantId = tenantId
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

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(_config.ExternalTimeoutSeconds));
            var response = await _httpClient.SendAsync(request, cts.Token);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Introspection endpoint returned {StatusCode}", response.StatusCode);
                return new ApiKeyAuthResult
                {
                    IsValid = false,
                    ErrorMessage = "API key validation failed"
                };
            }

            var content = await response.Content.ReadAsStringAsync(cts.Token);
            var result = JsonSerializer.Deserialize<IntrospectionResponse>(content, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (result?.Active == true)
            {
                _logger.LogDebug("Introspection validated key for tenant: {TenantId}", result.TenantId);
                return new ApiKeyAuthResult
                {
                    IsValid = true,
                    TenantId = result.TenantId
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

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(_config.ExternalTimeoutSeconds));
            var response = await _httpClient.SendAsync(request, cts.Token);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Custom endpoint returned {StatusCode}", response.StatusCode);
                return new ApiKeyAuthResult
                {
                    IsValid = false,
                    ErrorMessage = "API key validation failed"
                };
            }

            var content = await response.Content.ReadAsStringAsync(cts.Token);
            var result = JsonSerializer.Deserialize<CustomValidationResponse>(content, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (result?.Valid == true)
            {
                _logger.LogDebug("Custom endpoint validated key for tenant: {TenantId}", result.TenantId);
                return new ApiKeyAuthResult
                {
                    IsValid = true,
                    TenantId = result.TenantId
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
    private class IntrospectionResponse
    {
        public bool Active { get; init; }

        [JsonPropertyName("tenant_id")] public string? TenantId { get; init; }
    }

    /// <summary>
    ///     Response from custom validation endpoint
    /// </summary>
    private class CustomValidationResponse
    {
        public bool Valid { get; init; }

        [JsonPropertyName("tenant_id")] public string? TenantId { get; init; }

        public string? Error { get; init; }
    }
}

/// <summary>
///     Extension methods for adding API Key authentication middleware
/// </summary>
public static class ApiKeyAuthenticationExtensions
{
    /// <summary>
    ///     Adds API key authentication middleware to the application pipeline
    /// </summary>
    /// <param name="app">The application builder</param>
    /// <param name="config">API key configuration</param>
    /// <returns>The application builder for chaining</returns>
    public static IApplicationBuilder UseApiKeyAuthentication(
        this IApplicationBuilder app,
        ApiKeyConfig config)
    {
        if (!config.Enabled)
            return app;

        return app.UseMiddleware<ApiKeyAuthenticationMiddleware>(config);
    }
}