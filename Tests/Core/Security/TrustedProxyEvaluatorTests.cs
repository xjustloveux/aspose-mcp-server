using System.Net;
using AsposeMcpServer.Core.Security;

namespace AsposeMcpServer.Tests.Core.Security;

/// <summary>
///     Guards RB-09: gateway authentication modes read the caller's group and user identity from
///     request headers, which any client can set. The evaluator decides whether the immediate peer
///     is a gateway the operator declared trustworthy.
/// </summary>
public class TrustedProxyEvaluatorTests
{
    [Fact]
    public void NoConfiguredProxies_ShouldTrustNobody()
    {
        // A-04: an empty list means nothing has been trusted yet, not that everything is.
        // Trusting every peer here handed any direct caller the tenant identity headers that
        // gateway mode exists to protect.
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("203.0.113.9"), []));
    }

    [Fact]
    public void ExplicitAnyEntry_ShouldTrustAnyPeer()
    {
        // The deliberate escape hatch for a boundary enforced outside this process.
        Assert.True(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("203.0.113.9"),
            [TrustedProxyEvaluator.TrustAnyEntry]));
    }

    [Fact]
    public void ConfiguredProxies_ShouldRejectUnknownPeer()
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("203.0.113.9"), ["10.0.0.1"]));
    }

    [Fact]
    public void ConfiguredProxies_ShouldRejectUnknownRemoteAddress()
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(null, ["10.0.0.1"]));
    }

    [Theory]
    [InlineData("10.0.0.1", "10.0.0.1")]
    [InlineData("10.1.2.3", "10.0.0.0/8")]
    [InlineData("192.168.4.7", "192.168.4.0/24")]
    [InlineData("127.0.0.1", "127.0.0.0/8")]
    public void MatchingEntry_ShouldBeTrusted(string peer, string entry)
    {
        Assert.True(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse(peer), [entry]));
    }

    [Theory]
    [InlineData("11.0.0.1", "10.0.0.0/8")]
    [InlineData("192.168.5.7", "192.168.4.0/24")]
    [InlineData("10.0.0.2", "10.0.0.1")]
    public void NonMatchingEntry_ShouldBeRejected(string peer, string entry)
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse(peer), [entry]));
    }

    [Fact]
    public void IPv4MappedPeer_ShouldMatchIPv4Entry()
    {
        var mapped = IPAddress.Parse("10.0.0.5").MapToIPv6();

        Assert.True(TrustedProxyEvaluator.IsTrusted(mapped, ["10.0.0.0/8"]));
    }

    [Fact]
    public void IPv6Entry_ShouldMatchIPv6Peer()
    {
        Assert.True(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("2001:db8::5"), ["2001:db8::/32"]));
    }

    [Fact]
    public void MixedFamilies_ShouldNotMatch()
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("2001:db8::5"), ["10.0.0.0/8"]));
    }

    [Theory]
    [InlineData("not-an-address")]
    [InlineData("10.0.0.0/notanumber")]
    [InlineData("10.0.0.0/99")]
    [InlineData("   ")]
    public void MalformedEntry_ShouldNotTrust(string entry)
    {
        Assert.False(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("10.0.0.5"), [entry]));
    }

    [Fact]
    public void AnyMatchingEntryInList_ShouldTrust()
    {
        var trusted = new[] { "203.0.113.0/24", "10.0.0.0/8" };

        Assert.True(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("10.9.9.9"), trusted));
    }

    [Fact]
    public void ZeroLengthPrefix_ShouldMatchEveryAddressOfSameFamily()
    {
        Assert.True(TrustedProxyEvaluator.IsTrusted(IPAddress.Parse("203.0.113.9"), ["0.0.0.0/0"]));
    }
}
