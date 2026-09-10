namespace AsposeMcpServer.Core.Tracking;

/// <summary>
///     The one definition of "this request is the metrics endpoint" (A-07).
///     <para>
///         The authentication middleware used <c>StartsWithSegments(metricsPath)</c> and looked
///         only at whether anonymous metrics were allowed, while the tracking middleware served
///         metrics on an exact path match and only when metrics were enabled. Two predicates for
///         one endpoint means the request that skips authentication is not necessarily the request
///         that gets metrics: with metrics disabled, or with a path that is a prefix of a real
///         route, authentication could wave through something the metrics handler never claims.
///     </para>
/// </summary>
public static class MetricsRequest
{
    /// <summary>Whether this request is the configured metrics endpoint.</summary>
    /// <param name="path">The request path.</param>
    /// <param name="config">Tracking configuration, or null when tracking is not configured.</param>
    /// <returns><c>true</c> when metrics are enabled and the path is exactly the metrics path.</returns>
    public static bool Matches(PathString path, TrackingConfig? config)
    {
        if (config is not { MetricsEnabled: true }) return false;

        return path.Equals(config.MetricsPath, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Whether this request may skip authentication because it is the metrics endpoint and the
    ///     operator has opted into serving it anonymously.
    /// </summary>
    /// <param name="path">The request path.</param>
    /// <param name="config">Tracking configuration, or null when tracking is not configured.</param>
    /// <returns><c>true</c> only for an anonymous metrics request that will actually be served.</returns>
    public static bool AllowsAnonymousAccess(PathString path, TrackingConfig? config)
    {
        return Matches(path, config) && config is { MetricsRequireAuth: false };
    }
}
