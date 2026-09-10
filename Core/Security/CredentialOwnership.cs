using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;

namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Derives a session owner from the credential itself, for the authenticated requests that
///     carry no identity of their own.
///     <para>
///         Several success paths complete without producing either a group or a user: a JWT whose
///         configured group claim is absent and whose subject is empty, an introspection or custom
///         endpoint that answers "active" and nothing else, an API key configured with no group.
///         The identity built from those was indistinguishable from an unauthenticated one, so two
///         unrelated principals both landed in the shared anonymous bucket and could reach each
///         other's documents (R4-S04).
///     </para>
///     <para>
///         Refusing them instead would lock out deployments that authenticate but do not model
///         groups. A fingerprint of the credential keeps those working while still giving each
///         distinct credential its own bucket: it is stable for the life of the credential, and
///         different credentials never collide.
///     </para>
///     <para>
///         "For the life of the credential" is the whole of it, and it is why this is only ever
///         applied to an API key. A bearer token is reissued — on refresh, on re-login, on
///         rotation — and its bytes change each time, so fingerprinting one gave the same
///         principal a different bucket after every reissue: their earlier sessions became
///         unreachable, and the per-owner session quota reset as often as they cared to ask for
///         a new token (R5-S02). A token is bucketed by a claim that names the principal, and a
///         token that carries no such claim is refused.
///     </para>
/// </summary>
public static class CredentialOwnership
{
    /// <summary>Prefix that marks an owner id as derived rather than asserted by the caller.</summary>
    public const string Prefix = "credential:";

    /// <summary>
    ///     Claims that name the principal a token was issued for, rather than the token itself.
    ///     <para>
    ///         In order of preference: the subject, then the client an access token was issued to,
    ///         then the authorised party, then the two immutable identifiers Entra ID and several
    ///         other providers issue. Every one of them survives the token being reissued, which is
    ///         the property the bucket needs and the token's own bytes do not have.
    ///     </para>
    /// </summary>
    public static readonly string[] StablePrincipalClaims = ["sub", "client_id", "azp", "oid", "uid"];

    /// <summary>
    ///     Returns a stable, non-reversible owner id for a credential.
    /// </summary>
    /// <param name="credential">
    ///     The secret the caller authenticated with. It is hashed, never stored or logged; only
    ///     the first half of the digest is kept, which is far more than enough to keep distinct
    ///     credentials apart and leaves nothing to work backwards from.
    /// </param>
    /// <returns>An owner id, or <c>null</c> when there is no credential to derive one from.</returns>
    public static string? Fingerprint(string? credential)
    {
        if (string.IsNullOrEmpty(credential)) return null;

        var digest = SHA256.HashData(Encoding.UTF8.GetBytes(credential));
        return Prefix + Convert.ToHexString(digest, 0, 16).ToLowerInvariant();
    }

    /// <summary>
    ///     Returns the claim that names the principal a validated token belongs to.
    /// </summary>
    /// <param name="principal">The validated token's claims.</param>
    /// <returns>
    ///     The first stable principal claim the token carries, or <c>null</c> when it carries none —
    ///     in which case there is nothing to isolate this caller's sessions by, and the request has
    ///     to be refused rather than given a bucket that changes under it.
    /// </returns>
    public static string? StablePrincipal(ClaimsPrincipal principal)
    {
        foreach (var claim in StablePrincipalClaims)
        {
            var value = principal.FindFirst(claim)?.Value;
            if (!string.IsNullOrWhiteSpace(value)) return value;
        }

        return null;
    }
}
