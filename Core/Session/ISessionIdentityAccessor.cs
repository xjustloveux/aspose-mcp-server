namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Interface for accessing the current session identity from the request context
/// </summary>
public interface ISessionIdentityAccessor
{
    /// <summary>
    ///     Gets the identity of the current requestor
    /// </summary>
    /// <returns>The current session identity</returns>
    SessionIdentity GetCurrentIdentity();
}