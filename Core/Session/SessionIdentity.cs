using System.Text;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Represents the identity of a session owner or requestor
/// </summary>
public class SessionIdentity
{
    /// <summary>
    ///     Static anonymous identity instance (reused for efficiency)
    /// </summary>
    private static readonly SessionIdentity AnonymousInstance = new();

    /// <summary>
    ///     The tenant ID (organization/company identifier)
    /// </summary>
    public string? TenantId { get; init; }

    /// <summary>
    ///     The user ID within the tenant
    /// </summary>
    public string? UserId { get; init; }

    /// <summary>
    ///     Whether this identity is anonymous (no tenant or user info)
    /// </summary>
    public bool IsAnonymous => string.IsNullOrEmpty(TenantId) && string.IsNullOrEmpty(UserId);

    /// <summary>
    ///     Gets a static anonymous identity instance
    /// </summary>
    /// <returns>The shared anonymous identity instance</returns>
    public static SessionIdentity GetAnonymous()
    {
        return AnonymousInstance;
    }

    /// <summary>
    ///     Checks if this identity can access a session owned by the specified identity
    /// </summary>
    /// <param name="sessionOwner">The identity of the session owner</param>
    /// <param name="isolationMode">The current isolation mode</param>
    /// <returns>True if access is allowed</returns>
    public bool CanAccess(SessionIdentity sessionOwner, SessionIsolationMode isolationMode)
    {
        if (isolationMode == SessionIsolationMode.None)
            return true;

        if (IsAnonymous)
            return true;

        if (sessionOwner.IsAnonymous)
            return true;

        if (isolationMode == SessionIsolationMode.Tenant)
            return TenantId == sessionOwner.TenantId;

        return TenantId == sessionOwner.TenantId && UserId == sessionOwner.UserId;
    }

    /// <summary>
    ///     Gets the storage key for grouping sessions by owner.
    ///     IDs are Base64-encoded to prevent injection attacks from special characters.
    /// </summary>
    /// <param name="isolationMode">The current isolation mode</param>
    /// <returns>A key string for session storage grouping</returns>
    public string GetStorageKey(SessionIsolationMode isolationMode)
    {
        if (IsAnonymous || isolationMode == SessionIsolationMode.None)
            return "__anonymous__";

        // Encode IDs to prevent injection attacks from special characters like ":"
        var encodedTenantId = EncodeId(TenantId);

        if (isolationMode == SessionIsolationMode.Tenant)
            return $"tenant:{encodedTenantId}";

        var encodedUserId = EncodeId(UserId);
        return $"user:{encodedTenantId}:{encodedUserId}";
    }

    /// <summary>
    ///     Encodes an ID string using Base64 to prevent special character injection
    /// </summary>
    /// <param name="id">The ID to encode</param>
    /// <returns>Base64-encoded ID, or empty string if null/empty</returns>
    private static string EncodeId(string? id)
    {
        if (string.IsNullOrEmpty(id))
            return "";
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(id));
    }

    /// <summary>
    ///     Returns a string representation of this identity
    /// </summary>
    public override string ToString()
    {
        if (IsAnonymous)
            return "Anonymous";
        return $"Tenant={TenantId ?? "(null)"}, User={UserId ?? "(null)"}";
    }

    /// <summary>
    ///     Determines whether the specified object is equal to the current identity
    /// </summary>
    public override bool Equals(object? obj)
    {
        if (obj is not SessionIdentity other)
            return false;
        return TenantId == other.TenantId && UserId == other.UserId;
    }

    /// <summary>
    ///     Returns a hash code for this identity
    /// </summary>
    public override int GetHashCode()
    {
        return HashCode.Combine(TenantId, UserId);
    }
}