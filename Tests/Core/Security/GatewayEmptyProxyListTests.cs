using System.Net;
using AsposeMcpServer.Core.Security;

namespace AsposeMcpServer.Tests.Core.Security;

/// <summary>
///     Covers A-04. Gateway mode believes the caller's group and user headers, on the assumption
///     that only a trusted reverse proxy can reach the server. <see cref="TrustedProxyEvaluator" />
///     returned <c>true</c> for an empty proxy list, so an operator who enabled gateway mode
///     without configuring one got the opposite of what the mode promises: every direct caller
///     could name any tenant. Configuration validation only logged a warning, which nobody has to
///     read.
/// </summary>
public class GatewayEmptyProxyListTests
{
    [Fact]
    public void EmptyProxyList_ShouldNotTrustAnAddress()
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("203.0.113.7"), []));
    }

    [Fact]
    public void EmptyProxyList_ShouldNotTrustLoopbackEither()
    {
        // Loopback is not special here: a request arriving on it may still come from another
        // container or a port forward, and the mode's whole premise is a proxy in front.
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Loopback, []));
    }

    [Fact]
    public void ConfiguredProxy_ShouldStillBeTrusted()
    {
        Assert.True(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("10.0.0.5"), ["10.0.0.0/8"]));
    }

    [Fact]
    public void AddressOutsideTheConfiguredRange_ShouldStillBeRejected()
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("203.0.113.7"), ["10.0.0.0/8"]));
    }

    [Fact]
    public void UnknownAddress_ShouldBeRejected()
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(null, ["10.0.0.0/8"]));
    }

    [Fact]
    public void ExplicitAnySentinel_ShouldTrustEverything()
    {
        // The escape hatch for a deployment whose network boundary is enforced elsewhere. It has
        // to be spelled out, so choosing it is visible in the configuration rather than implied
        // by leaving a setting blank.
        Assert.True(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("203.0.113.7"), ["any"]));
        Assert.True(TrustedProxyEvaluator.IsTrusted(null, ["any"]));
    }
}
