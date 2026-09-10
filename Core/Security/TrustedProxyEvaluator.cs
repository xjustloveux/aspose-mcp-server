using System.Net;
using System.Net.Sockets;

namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Decides whether a request arrived from a proxy the operator declared trustworthy.
///     Gateway authentication modes read the caller's group and user identity from request headers,
///     which any client can set. Those headers are only meaningful when the request cannot reach the
///     server except through the gateway that sets them, so gateway mode consults this evaluator
///     before honouring them.
/// </summary>
public static class TrustedProxyEvaluator
{
    /// <summary>
    ///     The one entry that accepts every caller, for a deployment whose network boundary is
    ///     enforced outside this process. Spelled out so the choice is auditable.
    /// </summary>
    public const string TrustAnyEntry = "any";

    /// <summary>
    ///     Determines whether <paramref name="remoteAddress" /> matches one of the configured entries.
    /// </summary>
    /// <param name="remoteAddress">The immediate peer address of the request; may be null.</param>
    /// <param name="trustedProxies">
    ///     Configured entries, each a bare address (<c>10.0.0.5</c>), a CIDR range
    ///     (<c>10.0.0.0/8</c>), or the literal <c>any</c> to accept every caller.
    /// </param>
    /// <returns>
    ///     <c>true</c> when the address matches an entry, or when the list explicitly says
    ///     <c>any</c>; otherwise <c>false</c>, including for an empty list and an unknown address.
    /// </returns>
    /// <remarks>
    ///     An empty list used to mean "trust everyone", which inverted the guarantee gateway
    ///     mode exists to provide: an operator who turned the mode on but had not yet named a
    ///     proxy let any direct caller assert any tenant identity. The safe reading of "not
    ///     configured" is "nothing is trusted"; a deployment that really does enforce its
    ///     boundary elsewhere says so with the explicit <c>any</c> entry, which is visible in
    ///     the configuration instead of implied by a blank setting.
    /// </remarks>
    public static bool IsTrusted(IPAddress? remoteAddress, IReadOnlyList<string> trustedProxies)
    {
        if (trustedProxies.Count == 0) return false;

        if (trustedProxies.Any(entry =>
                entry.Trim().Equals(TrustAnyEntry, StringComparison.OrdinalIgnoreCase)))
            return true;

        if (remoteAddress == null) return false;

        var candidate = remoteAddress.IsIPv4MappedToIPv6 ? remoteAddress.MapToIPv4() : remoteAddress;

        return trustedProxies.Any(entry => Matches(candidate, entry));
    }

    /// <summary>
    ///     Matches one address against one configured entry.
    /// </summary>
    /// <param name="address">The normalised caller address.</param>
    /// <param name="entry">A bare address or a CIDR range.</param>
    /// <returns><c>true</c> when the address falls inside the entry.</returns>
    private static bool Matches(IPAddress address, string entry)
    {
        var text = entry.Trim();
        if (text.Length == 0) return false;

        var slash = text.IndexOf('/');
        if (slash < 0)
            return IPAddress.TryParse(text, out var single)
                   && Normalize(single).Equals(address);

        if (!IPAddress.TryParse(text[..slash], out var network)) return false;
        if (!int.TryParse(text[(slash + 1)..], out var prefixLength)) return false;

        network = Normalize(network);
        if (network.AddressFamily != address.AddressFamily) return false;

        var maxBits = network.AddressFamily == AddressFamily.InterNetwork ? 32 : 128;
        if (prefixLength < 0 || prefixLength > maxBits) return false;

        var networkBytes = network.GetAddressBytes();
        var addressBytes = address.GetAddressBytes();
        var fullBytes = prefixLength / 8;
        var remainingBits = prefixLength % 8;

        for (var i = 0; i < fullBytes; i++)
            if (networkBytes[i] != addressBytes[i])
                return false;

        if (remainingBits == 0) return true;

        var mask = (byte)(0xFF << (8 - remainingBits));
        return (networkBytes[fullBytes] & mask) == (addressBytes[fullBytes] & mask);
    }

    /// <summary>
    ///     Collapses an IPv4-mapped IPv6 address to its IPv4 form so both notations compare equal.
    /// </summary>
    /// <param name="address">The address to normalise.</param>
    /// <returns>The normalised address.</returns>
    private static IPAddress Normalize(IPAddress address)
    {
        return address.IsIPv4MappedToIPv6 ? address.MapToIPv4() : address;
    }
}
