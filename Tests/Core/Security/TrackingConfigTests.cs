using AsposeMcpServer.Core.Security;

namespace AsposeMcpServer.Tests.Core.Security;

public class TrackingConfigTests
{
    [Fact]
    public void TrackingConfig_DefaultValues_ShouldBeCorrect()
    {
        var config = new TrackingConfig();

        Assert.True(config.LogEnabled);
        Assert.Single(config.LogTargets);
        Assert.Equal(LogTarget.Console, config.LogTargets[0]);
        Assert.False(config.WebhookEnabled);
        Assert.Equal(5, config.WebhookTimeoutSeconds);
        Assert.False(config.MetricsEnabled);
        Assert.Equal("/metrics", config.MetricsPath);
    }

    [Fact]
    public void TrackingConfig_LoadFromEnvironment_WithLogSettings()
    {
        Environment.SetEnvironmentVariable("ASPOSE_LOG_ENABLED", "true");
        Environment.SetEnvironmentVariable("ASPOSE_LOG_TARGETS", "Console,EventLog");

        try
        {
            var config = TrackingConfig.LoadFromArgs([]);
            Assert.True(config.LogEnabled);
            Assert.Equal(2, config.LogTargets.Length);
            Assert.Contains(LogTarget.Console, config.LogTargets);
            Assert.Contains(LogTarget.EventLog, config.LogTargets);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_LOG_ENABLED", null);
            Environment.SetEnvironmentVariable("ASPOSE_LOG_TARGETS", null);
        }
    }

    [Fact]
    public void TrackingConfig_LoadFromEnvironment_WithWebhookSettings()
    {
        Environment.SetEnvironmentVariable("ASPOSE_WEBHOOK_URL", "https://example.com/webhook");
        Environment.SetEnvironmentVariable("ASPOSE_WEBHOOK_AUTH_HEADER", "Bearer token123");
        Environment.SetEnvironmentVariable("ASPOSE_WEBHOOK_TIMEOUT", "10");

        try
        {
            var config = TrackingConfig.LoadFromArgs([]);
            Assert.True(config.WebhookEnabled); // Auto-enabled when URL is set
            Assert.Equal("https://example.com/webhook", config.WebhookUrl);
            Assert.Equal("Bearer token123", config.WebhookAuthHeader);
            Assert.Equal(10, config.WebhookTimeoutSeconds);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_WEBHOOK_URL", null);
            Environment.SetEnvironmentVariable("ASPOSE_WEBHOOK_AUTH_HEADER", null);
            Environment.SetEnvironmentVariable("ASPOSE_WEBHOOK_TIMEOUT", null);
        }
    }

    [Fact]
    public void TrackingConfig_LoadFromEnvironment_WithMetricsSettings()
    {
        Environment.SetEnvironmentVariable("ASPOSE_METRICS_ENABLED", "true");
        Environment.SetEnvironmentVariable("ASPOSE_METRICS_PATH", "/custom-metrics");

        try
        {
            var config = TrackingConfig.LoadFromArgs([]);
            Assert.True(config.MetricsEnabled);
            Assert.Equal("/custom-metrics", config.MetricsPath);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_METRICS_ENABLED", null);
            Environment.SetEnvironmentVariable("ASPOSE_METRICS_PATH", null);
        }
    }

    [Fact]
    public void TrackingEvent_DefaultValues_ShouldHaveTimestamp()
    {
        var evt = new TrackingEvent();

        Assert.True(evt.Timestamp > DateTime.MinValue);
        Assert.False(evt.Success); // Default is false
        Assert.Equal(0, evt.DurationMs);
    }

    [Fact]
    public void TrackingEvent_ShouldHoldAllProperties()
    {
        var evt = new TrackingEvent
        {
            TenantId = "tenant1",
            UserId = "user1",
            Tool = "pdf_text",
            Operation = "get",
            DurationMs = 150,
            Success = true,
            SessionMemoryMb = 45.2,
            SessionId = "sess_123",
            RequestId = "req_abc"
        };

        Assert.Equal("tenant1", evt.TenantId);
        Assert.Equal("user1", evt.UserId);
        Assert.Equal("pdf_text", evt.Tool);
        Assert.Equal("get", evt.Operation);
        Assert.Equal(150, evt.DurationMs);
        Assert.True(evt.Success);
        Assert.Equal(45.2, evt.SessionMemoryMb);
        Assert.Equal("sess_123", evt.SessionId);
        Assert.Equal("req_abc", evt.RequestId);
    }
}