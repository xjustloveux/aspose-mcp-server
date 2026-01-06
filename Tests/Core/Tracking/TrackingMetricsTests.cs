using AsposeMcpServer.Core.Tracking;

namespace AsposeMcpServer.Tests.Core.Tracking;

/// <summary>
///     Unit tests for TrackingMetrics class
/// </summary>
public class TrackingMetricsTests
{
    #region Edge Cases

    [Fact]
    public void GetPrometheusMetrics_ToolWithSpecialChars_ShouldBeSanitized()
    {
        var metrics = new TrackingMetrics();
        metrics.RecordRequest(new TrackingEvent { Tool = "tool\"with\"quotes", Success = true });
        metrics.RecordRequest(new TrackingEvent { Tool = "tool\\with\\backslash", Success = true });
        metrics.RecordRequest(new TrackingEvent { Tool = "tool\nwith\nnewline", Success = true });

        var output = metrics.GetPrometheusMetrics();

        Assert.Contains("tool=\"tool\\\"with\\\"quotes\"", output);
        Assert.Contains("tool=\"tool\\\\with\\\\backslash\"", output);
        Assert.Contains("tool=\"tool\\nwith\\nnewline\"", output);
    }

    #endregion

    #region RecordRequest Tests

    [Fact]
    public void RecordRequest_ShouldIncrementTotalRequests()
    {
        var metrics = new TrackingMetrics();
        var evt = new TrackingEvent
        {
            Tool = "pdf_text",
            Success = true,
            DurationMs = 100
        };
        metrics.RecordRequest(evt);
        metrics.RecordRequest(evt);
        var output = metrics.GetPrometheusMetrics();
        Assert.Contains("aspose_mcp_requests_total 2", output);
    }

    [Fact]
    public void RecordRequest_ShouldTrackSuccessfulRequests()
    {
        var metrics = new TrackingMetrics();
        var successEvt = new TrackingEvent { Success = true };
        var failEvt = new TrackingEvent { Success = false };
        metrics.RecordRequest(successEvt);
        metrics.RecordRequest(successEvt);
        metrics.RecordRequest(failEvt);
        var output = metrics.GetPrometheusMetrics();
        Assert.Contains("aspose_mcp_requests_successful_total 2", output);
        Assert.Contains("aspose_mcp_requests_failed_total 1", output);
    }

    [Fact]
    public void RecordRequest_ShouldTrackByTool()
    {
        var metrics = new TrackingMetrics();
        metrics.RecordRequest(new TrackingEvent { Tool = "pdf_text", Success = true });
        metrics.RecordRequest(new TrackingEvent { Tool = "pdf_text", Success = true });
        metrics.RecordRequest(new TrackingEvent { Tool = "word_file", Success = true });
        var output = metrics.GetPrometheusMetrics();
        Assert.Contains("aspose_mcp_requests_by_tool{tool=\"pdf_text\"} 2", output);
        Assert.Contains("aspose_mcp_requests_by_tool{tool=\"word_file\"} 1", output);
    }

    [Fact]
    public void RecordRequest_ShouldCalculateAverageDuration()
    {
        var metrics = new TrackingMetrics();
        metrics.RecordRequest(new TrackingEvent { DurationMs = 100, Success = true });
        metrics.RecordRequest(new TrackingEvent { DurationMs = 200, Success = true });
        metrics.RecordRequest(new TrackingEvent { DurationMs = 300, Success = true });
        var output = metrics.GetPrometheusMetrics();
        // Average should be (100+200+300)/3 = 200
        Assert.Contains("aspose_mcp_request_duration_ms_avg 200", output);
    }

    [Fact]
    public void GetPrometheusMetrics_ShouldIncludeMemoryMetric()
    {
        var metrics = new TrackingMetrics();
        var output = metrics.GetPrometheusMetrics();
        Assert.Contains("aspose_mcp_memory_mb", output);
    }

    [Fact]
    public void GetPrometheusMetrics_ShouldIncludeHelpAndTypeLines()
    {
        var metrics = new TrackingMetrics();
        metrics.RecordRequest(new TrackingEvent { Tool = "test", Success = true });
        var output = metrics.GetPrometheusMetrics();
        Assert.Contains("# HELP aspose_mcp_requests_total", output);
        Assert.Contains("# TYPE aspose_mcp_requests_total counter", output);
        Assert.Contains("# HELP aspose_mcp_requests_successful_total", output);
        Assert.Contains("# TYPE aspose_mcp_requests_successful_total counter", output);
    }

    [Fact]
    public void RecordRequest_NullTool_ShouldNotTrackByTool()
    {
        var metrics = new TrackingMetrics();
        metrics.RecordRequest(new TrackingEvent { Tool = null, Success = true });
        var output = metrics.GetPrometheusMetrics();
        Assert.DoesNotContain("aspose_mcp_requests_by_tool{tool=\"\"}", output);
    }

    #endregion

    #region GetPrometheusMetrics Tests

    [Fact]
    public void GetPrometheusMetrics_ZeroRequests_ShouldReturnZeroAverage()
    {
        var metrics = new TrackingMetrics();
        var output = metrics.GetPrometheusMetrics();
        Assert.Contains("aspose_mcp_request_duration_ms_avg 0", output);
    }

    [Fact]
    public async Task GetPrometheusMetrics_ShouldBeThreadSafe()
    {
        var metrics = new TrackingMetrics();
        List<Task> tasks = [];

        for (var i = 0; i < 100; i++)
        {
            var index = i; // Capture loop variable to avoid closure issue
            var toolName = $"tool_{index % 5}";
            tasks.Add(Task.Run(() => metrics.RecordRequest(new TrackingEvent
            {
                Tool = toolName,
                Success = index % 2 == 0,
                DurationMs = index * 10
            })));
        }

        await Task.WhenAll(tasks);
        var output = metrics.GetPrometheusMetrics();
        Assert.Contains("aspose_mcp_requests_total 100", output);
    }

    #endregion
}
