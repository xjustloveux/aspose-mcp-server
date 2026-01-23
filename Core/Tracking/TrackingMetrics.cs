using System.Diagnostics;
using System.Text;

namespace AsposeMcpServer.Core.Tracking;

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
            sb.AppendLine("# HELP aspose_mcp_requests_total Total number of MCP requests");
            sb.AppendLine("# TYPE aspose_mcp_requests_total counter");
            sb.AppendLine($"aspose_mcp_requests_total {_totalRequests}");

            sb.AppendLine("# HELP aspose_mcp_requests_successful_total Total number of successful MCP requests");
            sb.AppendLine("# TYPE aspose_mcp_requests_successful_total counter");
            sb.AppendLine($"aspose_mcp_requests_successful_total {_successfulRequests}");

            sb.AppendLine("# HELP aspose_mcp_requests_failed_total Total number of failed MCP requests");
            sb.AppendLine("# TYPE aspose_mcp_requests_failed_total counter");
            sb.AppendLine($"aspose_mcp_requests_failed_total {_failedRequests}");

            var avgDuration = _totalRequests > 0 ? (double)_totalDurationMs / _totalRequests : 0;
            sb.AppendLine("# HELP aspose_mcp_request_duration_ms_avg Average request duration in milliseconds");
            sb.AppendLine("# TYPE aspose_mcp_request_duration_ms_avg gauge");
            sb.AppendLine($"aspose_mcp_request_duration_ms_avg {avgDuration:F2}");

            sb.AppendLine("# HELP aspose_mcp_requests_by_tool Total requests by tool");
            sb.AppendLine("# TYPE aspose_mcp_requests_by_tool counter");
            foreach (var (tool, count) in _requestsByTool)
                sb.AppendLine($"aspose_mcp_requests_by_tool{{tool=\"{SanitizeLabelValue(tool)}\"}} {count}");

            var memoryMb = Process.GetCurrentProcess().WorkingSet64 / (1024.0 * 1024.0);
            sb.AppendLine("# HELP aspose_mcp_memory_mb Current memory usage in MB");
            sb.AppendLine("# TYPE aspose_mcp_memory_mb gauge");
            sb.AppendLine($"aspose_mcp_memory_mb {memoryMb:F2}");
        }

        return sb.ToString();
    }

    /// <summary>
    ///     Sanitizes a string for use as a Prometheus label value.
    ///     Escapes backslashes, double quotes, and newlines.
    /// </summary>
    /// <param name="value">The label value to sanitize</param>
    /// <returns>Sanitized label value safe for Prometheus format</returns>
    private static string SanitizeLabelValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        return value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n");
    }
}
