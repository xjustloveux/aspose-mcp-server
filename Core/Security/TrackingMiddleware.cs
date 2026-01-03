using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Middleware for tracking tool invocations and outputting to various targets
/// </summary>
public class TrackingMiddleware
{
    /// <summary>
    ///     Tracking configuration
    /// </summary>
    private readonly TrackingConfig _config;

    /// <summary>
    ///     HTTP client for webhook calls
    /// </summary>
    private readonly HttpClient _httpClient;

    /// <summary>
    ///     Logger for tracking events
    /// </summary>
    private readonly ILogger<TrackingMiddleware> _logger;

    /// <summary>
    ///     Metrics collector for Prometheus format
    /// </summary>
    private readonly TrackingMetrics _metrics;

    /// <summary>
    ///     Next middleware in the pipeline
    /// </summary>
    private readonly RequestDelegate _next;

    /// <summary>
    ///     Creates a new tracking middleware instance
    /// </summary>
    /// <param name="next">Next middleware delegate</param>
    /// <param name="config">Tracking configuration</param>
    /// <param name="logger">Logger instance</param>
    /// <param name="httpClientFactory">Optional HTTP client factory</param>
    public TrackingMiddleware(
        RequestDelegate next,
        TrackingConfig config,
        ILogger<TrackingMiddleware> logger,
        IHttpClientFactory? httpClientFactory = null)
    {
        _next = next;
        _config = config;
        _logger = logger;
        _httpClient = httpClientFactory?.CreateClient("Tracking") ?? new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(config.WebhookTimeoutSeconds)
        };
        _metrics = new TrackingMetrics();
    }

    /// <summary>
    ///     Processes an HTTP request to track tool invocations
    /// </summary>
    /// <param name="context">HTTP context for the current request</param>
    public async Task InvokeAsync(HttpContext context)
    {
        if (_config.MetricsEnabled &&
            context.Request.Path.Equals(_config.MetricsPath, StringComparison.OrdinalIgnoreCase))
        {
            await HandleMetricsRequest(context);
            return;
        }

        var requestId = Guid.NewGuid().ToString("N")[..12];
        context.Items["RequestId"] = requestId;

        var stopwatch = Stopwatch.StartNew();
        string? error = null;
        var success = true;

        try
        {
            await _next(context);

            if (context.Response.StatusCode >= 400)
            {
                success = false;
                error = $"HTTP {context.Response.StatusCode}";
            }
        }
        catch (Exception ex)
        {
            success = false;
            error = ex.Message;
            throw;
        }
        finally
        {
            stopwatch.Stop();

            var trackingEvent = BuildTrackingEvent(context, stopwatch.ElapsedMilliseconds, success, error, requestId);

            if (trackingEvent.Tool != null) await TrackEventAsync(trackingEvent);
        }
    }

    /// <summary>
    ///     Builds a tracking event from the HTTP context
    /// </summary>
    /// <param name="context">HTTP context</param>
    /// <param name="durationMs">Request duration in milliseconds</param>
    /// <param name="success">Whether the request succeeded</param>
    /// <param name="error">Error message if failed</param>
    /// <param name="requestId">Correlation request ID</param>
    /// <returns>Tracking event with populated fields</returns>
    private TrackingEvent BuildTrackingEvent(HttpContext context, long durationMs, bool success, string? error,
        string requestId)
    {
        var tool = context.Items["ToolName"]?.ToString();
        var operation = context.Items["ToolOperation"]?.ToString();
        var sessionId = context.Items["SessionId"]?.ToString();

        if (string.IsNullOrEmpty(tool))
        {
            var path = context.Request.Path.Value ?? "";
            if (path.Contains("/sse") || path.Contains("/ws") || path.Contains("/mcp")) tool = "mcp_request";
        }

        return new TrackingEvent
        {
            Timestamp = DateTime.UtcNow,
            TenantId = context.Items["TenantId"]?.ToString(),
            UserId = context.Items["UserId"]?.ToString(),
            Tool = tool,
            Operation = operation,
            DurationMs = durationMs,
            Success = success,
            Error = error,
            SessionMemoryMb = GetSessionMemoryMb(),
            SessionId = sessionId,
            RequestId = requestId
        };
    }

    /// <summary>
    ///     Gets the current process memory usage in megabytes
    /// </summary>
    /// <returns>Memory usage in MB</returns>
    private double GetSessionMemoryMb()
    {
        var process = Process.GetCurrentProcess();
        return process.WorkingSet64 / (1024.0 * 1024.0);
    }

    /// <summary>
    ///     Tracks an event to all configured outputs
    /// </summary>
    /// <param name="trackingEvent">Event to track</param>
    /// <returns>Completed task</returns>
    private Task TrackEventAsync(TrackingEvent trackingEvent)
    {
        _metrics.RecordRequest(trackingEvent);

        if (_config.LogEnabled) LogEvent(trackingEvent);

        if (_config.WebhookEnabled && !string.IsNullOrEmpty(_config.WebhookUrl)) _ = SendWebhookAsync(trackingEvent);

        return Task.CompletedTask;
    }

    /// <summary>
    ///     Logs an event to configured log targets
    /// </summary>
    /// <param name="trackingEvent">Event to log</param>
    private void LogEvent(TrackingEvent trackingEvent)
    {
        var json = JsonSerializer.Serialize(trackingEvent, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = false
        });

        foreach (var target in _config.LogTargets)
            switch (target)
            {
                case LogTarget.Console:
                    if (trackingEvent.Success)
                        _logger.LogInformation("Tracking: {Event}", json);
                    else
                        _logger.LogWarning("Tracking: {Event}", json);
                    break;

                case LogTarget.EventLog:
                    if (OperatingSystem.IsWindows()) WriteToEventLog(trackingEvent, json);
                    break;
            }
    }

    /// <summary>
    ///     Writes an event to Windows Event Log (Windows only)
    /// </summary>
    /// <param name="trackingEvent">Event to write</param>
    /// <param name="json">JSON-serialized event data</param>
    // ReSharper disable UnusedParameter.Local
    private void WriteToEventLog(TrackingEvent trackingEvent, string json)
        // ReSharper restore UnusedParameter.Local
    {
#if WINDOWS
        try
        {
            var source = "AsposeMcpServer";
            var logName = "Application";

            if (!System.Diagnostics.EventLog.SourceExists(source))
            {
                System.Diagnostics.EventLog.CreateEventSource(source, logName);
            }

            var entryType = trackingEvent.Success
                ? System.Diagnostics.EventLogEntryType.Information
                : System.Diagnostics.EventLogEntryType.Warning;

            System.Diagnostics.EventLog.WriteEntry(source, json, entryType);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to write to Event Log");
        }
#endif
    }

    /// <summary>
    ///     Sends an event to the configured webhook endpoint
    /// </summary>
    /// <param name="trackingEvent">Event to send</param>
    private async Task SendWebhookAsync(TrackingEvent trackingEvent)
    {
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, _config.WebhookUrl);
            request.Content = new StringContent(
                JsonSerializer.Serialize(trackingEvent, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                }),
                Encoding.UTF8,
                "application/json");

            if (!string.IsNullOrEmpty(_config.WebhookAuthHeader))
                request.Headers.TryAddWithoutValidation("Authorization", _config.WebhookAuthHeader);

            await _httpClient.SendAsync(request);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to send webhook");
        }
    }

    /// <summary>
    ///     Handles Prometheus metrics endpoint requests
    /// </summary>
    /// <param name="context">HTTP context</param>
    private async Task HandleMetricsRequest(HttpContext context)
    {
        context.Response.ContentType = "text/plain; version=0.0.4; charset=utf-8";
        await context.Response.WriteAsync(_metrics.GetPrometheusMetrics());
    }
}

/// <summary>
///     Simple in-memory metrics collector for Prometheus format
/// </summary>
public class TrackingMetrics
{
    /// <summary>
    ///     Lock for thread-safe metric updates
    /// </summary>
    private readonly object _lock = new();

    /// <summary>
    ///     Request counts by tool name
    /// </summary>
    private readonly Dictionary<string, long> _requestsByTool = new();

    /// <summary>
    ///     Total failed request count
    /// </summary>
    private long _failedRequests;

    /// <summary>
    ///     Total successful request count
    /// </summary>
    private long _successfulRequests;

    /// <summary>
    ///     Total cumulative request duration in milliseconds
    /// </summary>
    private long _totalDurationMs;

    /// <summary>
    ///     Total request count
    /// </summary>
    private long _totalRequests;

    /// <summary>
    ///     Records a tracking event for metrics collection
    /// </summary>
    /// <param name="evt">The tracking event to record</param>
    public void RecordRequest(TrackingEvent evt)
    {
        lock (_lock)
        {
            _totalRequests++;
            _totalDurationMs += evt.DurationMs;

            if (evt.Success)
                _successfulRequests++;
            else
                _failedRequests++;

            if (!string.IsNullOrEmpty(evt.Tool))
            {
                _requestsByTool.TryGetValue(evt.Tool, out var count);
                _requestsByTool[evt.Tool] = count + 1;
            }
        }
    }

    /// <summary>
    ///     Returns metrics in Prometheus text format
    /// </summary>
    /// <returns>Prometheus-formatted metrics string</returns>
    public string GetPrometheusMetrics()
    {
        var sb = new StringBuilder();

        lock (_lock)
        {
            // Total requests
            sb.AppendLine("# HELP aspose_mcp_requests_total Total number of MCP requests");
            sb.AppendLine("# TYPE aspose_mcp_requests_total counter");
            sb.AppendLine($"aspose_mcp_requests_total {_totalRequests}");

            // Successful requests
            sb.AppendLine("# HELP aspose_mcp_requests_successful_total Total number of successful MCP requests");
            sb.AppendLine("# TYPE aspose_mcp_requests_successful_total counter");
            sb.AppendLine($"aspose_mcp_requests_successful_total {_successfulRequests}");

            // Failed requests
            sb.AppendLine("# HELP aspose_mcp_requests_failed_total Total number of failed MCP requests");
            sb.AppendLine("# TYPE aspose_mcp_requests_failed_total counter");
            sb.AppendLine($"aspose_mcp_requests_failed_total {_failedRequests}");

            // Average duration
            var avgDuration = _totalRequests > 0 ? (double)_totalDurationMs / _totalRequests : 0;
            sb.AppendLine("# HELP aspose_mcp_request_duration_ms_avg Average request duration in milliseconds");
            sb.AppendLine("# TYPE aspose_mcp_request_duration_ms_avg gauge");
            sb.AppendLine($"aspose_mcp_request_duration_ms_avg {avgDuration:F2}");

            // Requests by tool
            sb.AppendLine("# HELP aspose_mcp_requests_by_tool Total requests by tool");
            sb.AppendLine("# TYPE aspose_mcp_requests_by_tool counter");
            foreach (var (tool, count) in _requestsByTool)
                sb.AppendLine($"aspose_mcp_requests_by_tool{{tool=\"{tool}\"}} {count}");

            // Memory usage
            var memoryMb = Process.GetCurrentProcess().WorkingSet64 / (1024.0 * 1024.0);
            sb.AppendLine("# HELP aspose_mcp_memory_mb Current memory usage in MB");
            sb.AppendLine("# TYPE aspose_mcp_memory_mb gauge");
            sb.AppendLine($"aspose_mcp_memory_mb {memoryMb:F2}");
        }

        return sb.ToString();
    }
}

/// <summary>
///     Extension methods for adding tracking middleware
/// </summary>
public static class TrackingExtensions
{
    /// <summary>
    ///     Adds tracking middleware to the application pipeline
    /// </summary>
    /// <param name="app">The application builder</param>
    /// <param name="config">Tracking configuration</param>
    /// <returns>The application builder for chaining</returns>
    public static IApplicationBuilder UseTracking(
        this IApplicationBuilder app,
        TrackingConfig config)
    {
        if (config is { LogEnabled: false, WebhookEnabled: false, MetricsEnabled: false })
            return app;

        return app.UseMiddleware<TrackingMiddleware>(config);
    }
}