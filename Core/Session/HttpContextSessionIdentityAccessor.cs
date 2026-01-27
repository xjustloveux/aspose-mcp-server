namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Gets session identity from HttpContext (for HTTP/WebSocket modes)
/// </summary>
public class HttpContextSessionIdentityAccessor : ISessionIdentityAccessor
{
    /// <summary>
    ///     The HTTP context accessor for retrieving the current HTTP context.
    /// </summary>
    private readonly IHttpContextAccessor _httpContextAccessor;

    /// <summary>
    ///     Initializes a new instance of HttpContextSessionIdentityAccessor
    /// </summary>
    /// <param name="httpContextAccessor">The HTTP context accessor</param>
    public HttpContextSessionIdentityAccessor(IHttpContextAccessor httpContextAccessor)
    {
        _httpContextAccessor = httpContextAccessor;
    }

    /// <inheritdoc />
    public SessionIdentity GetCurrentIdentity()
    {
        var context = _httpContextAccessor.HttpContext;
        if (context == null)
            return SessionIdentity.GetAnonymous();

        var groupId = context.Items["GroupId"]?.ToString();
        var userId = context.Items["UserId"]?.ToString();

        return new SessionIdentity { GroupId = groupId, UserId = userId };
    }
}
