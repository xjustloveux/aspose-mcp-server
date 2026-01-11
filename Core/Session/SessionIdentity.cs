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
    ///     The group identifier for session isolation.
    ///     This value is determined by the GROUP_IDENTIFIER_CLAIM configuration
    ///     (e.g., tenant_id, team_id, sub, or any custom claim).
    /// </summary>
    public string? GroupId { get; init; }

    /// <summary>
    ///     The user ID for audit and logging purposes
    /// </summary>
    public string? UserId { get; init; }

    /// <summary>
    ///     Whether this identity is anonymous (no group or user info)
    /// </summary>
    public bool IsAnonymous => string.IsNullOrEmpty(GroupId) && string.IsNullOrEmpty(UserId);

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

        return GroupId == sessionOwner.GroupId;
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

        var encodedGroupId = EncodeId(GroupId);
        return $"group:{encodedGroupId}";
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
    /// <returns>A human-readable string representation</returns>
    public override string ToString()
    {
        if (IsAnonymous)
            return "Anonymous";
        return $"Group={GroupId ?? "(null)"}, User={UserId ?? "(null)"}";
    }

    /// <summary>
    ///     Determines whether the specified object is equal to the current identity
    /// </summary>
    /// <param name="obj">The object to compare with the current identity</param>
    /// <returns>True if the specified object is equal to the current identity</returns>
    public override bool Equals(object? obj)
    {
        if (obj is not SessionIdentity other)
            return false;
        return GroupId == other.GroupId && UserId == other.UserId;
    }

    /// <summary>
    ///     Returns a hash code for this identity
    /// </summary>
    /// <returns>A hash code for the current identity</returns>
    public override int GetHashCode()
    {
        return HashCode.Combine(GroupId, UserId);
    }
}
