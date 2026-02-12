using AsposeMcpServer.Core.Tracking;

namespace AsposeMcpServer.Tests.Core.Tracking;

/// <summary>
///     Unit tests for TrackingExtensions class
/// </summary>
public class TrackingExtensionsTests
{
    [Fact]
    public void UseTracking_WithAllDisabled_ShouldNotAddMiddleware()
    {
        var config = new TrackingConfig
        {
            LogEnabled = false,
            WebhookEnabled = false,
            MetricsEnabled = false
        };

        Assert.False(config.LogEnabled);
        Assert.False(config.WebhookEnabled);
        Assert.False(config.MetricsEnabled);
    }

    [Fact]
    public void TrackingConfig_DefaultValues_ShouldBeCorrect()
    {
        var config = new TrackingConfig();

        Assert.True(config.LogEnabled);
        Assert.Contains(LogTarget.Console, config.LogTargets);
        Assert.False(config.WebhookEnabled);
        Assert.False(config.MetricsEnabled);
        Assert.Equal("/metrics", config.MetricsPath);
        Assert.Equal(5, config.WebhookTimeoutSeconds);
    }
}
