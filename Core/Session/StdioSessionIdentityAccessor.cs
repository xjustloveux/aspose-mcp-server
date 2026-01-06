namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Gets session identity for Stdio mode.
///     Reads from environment variables if available (for WebSocket child processes),
///     otherwise returns anonymous identity.
/// </summary>
public class StdioSessionIdentityAccessor : ISessionIdentityAccessor
{
    /// <summary>
    ///     Cached identity instance (environment variables don't change during process lifetime)
    /// </summary>
    private readonly SessionIdentity _identity;

    /// <summary>
    ///     Initializes the accessor by reading identity from environment variables
    /// </summary>
    public StdioSessionIdentityAccessor()
    {
        var tenantId = Environment.GetEnvironmentVariable("ASPOSE_SESSION_TENANT_ID");
        var userId = Environment.GetEnvironmentVariable("ASPOSE_SESSION_USER_ID");

        if (string.IsNullOrEmpty(tenantId) && string.IsNullOrEmpty(userId))
            _identity = SessionIdentity.GetAnonymous();
        else
            _identity = new SessionIdentity
            {
                TenantId = tenantId,
                UserId = userId
            };
    }

    /// <inheritdoc />
    public SessionIdentity GetCurrentIdentity()
    {
        return _identity;
    }
}