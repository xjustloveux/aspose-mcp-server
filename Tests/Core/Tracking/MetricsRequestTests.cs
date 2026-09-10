using AsposeMcpServer.Core.Tracking;
using Microsoft.AspNetCore.Http;

namespace AsposeMcpServer.Tests.Core.Tracking;

/// <summary>
///     Covers A-07. Authentication decided "this is metrics" with
///     <c>StartsWithSegments(metricsPath)</c> and only consulted whether anonymous access was
///     allowed; the middleware that actually serves metrics used an exact match and also required
///     metrics to be enabled. Two definitions of one endpoint meant the request that skipped
///     authentication was not necessarily the request that would be served.
/// </summary>
public class MetricsRequestTests
{
    /// <summary>Builds a tracking configuration for the metrics endpoint.</summary>
    /// <param name="enabled">Whether metrics are served at all.</param>
    /// <param name="requireAuth">Whether the endpoint requires authentication.</param>
    /// <param name="path">The configured metrics path.</param>
    /// <returns>The configuration.</returns>
    private static TrackingConfig Config(bool enabled, bool requireAuth, string path = "/metrics")
    {
        return new TrackingConfig
        {
            MetricsEnabled = enabled,
            MetricsRequireAuth = requireAuth,
            MetricsPath = path
        };
    }

    [Fact]
    public void ExactPath_WithMetricsEnabledAndAnonymousAllowed_ShouldBeExempt()
    {
        Assert.True(MetricsRequest.AllowsAnonymousAccess(new PathString("/metrics"),
            Config(true, false)));
    }

    [Fact]
    public void MetricsDisabled_ShouldNotBeExemptEvenWhenAnonymousIsConfigured()
    {
        // Nothing serves metrics here, so exempting the path only removes authentication from it.
        Assert.False(MetricsRequest.AllowsAnonymousAccess(new PathString("/metrics"),
            Config(false, false)));
    }

    [Fact]
    public void ChildPath_ShouldNotBeExempt()
    {
        // The prefix match used to cover every path underneath the metrics path.
        Assert.False(MetricsRequest.AllowsAnonymousAccess(new PathString("/metrics/../mcp"),
            Config(true, false)));
        Assert.False(MetricsRequest.AllowsAnonymousAccess(new PathString("/metrics/sub"),
            Config(true, false)));
    }

    [Fact]
    public void AuthRequired_ShouldNotBeExempt()
    {
        Assert.False(MetricsRequest.AllowsAnonymousAccess(new PathString("/metrics"),
            Config(true, true)));
    }

    [Fact]
    public void NoTrackingConfiguration_ShouldNotBeExempt()
    {
        Assert.False(MetricsRequest.AllowsAnonymousAccess(new PathString("/metrics"), null));
    }

    [Fact]
    public void ServingAndAuthentication_ShouldAgree()
    {
        // The property that was missing: one predicate, so the request that skips authentication
        // is exactly the request that gets served.
        var config = Config(true, false);
        var path = new PathString("/metrics");

        Assert.Equal(MetricsRequest.Matches(path, config),
            MetricsRequest.AllowsAnonymousAccess(path, config));
    }

    [Theory]
    [InlineData("/mcp")]
    [InlineData("/health")]
    [InlineData("/")]
    public void MetricsPathCollidingWithAReservedRoute_ShouldBeRefused(string path)
    {
        var config = Config(true, false, path);

        Assert.Throws<InvalidOperationException>(() => config.Validate());
    }

    [Fact]
    public void OrdinaryMetricsPath_ShouldValidate()
    {
        var config = Config(true, false, "/internal-metrics");

        config.Validate();

        Assert.Equal("/internal-metrics", config.MetricsPath);
    }
}
