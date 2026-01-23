namespace AsposeMcpServer.Core.Security;

/// <summary>
///     JWT authentication verification mode
/// </summary>
public enum JwtMode
{
    /// <summary>
    ///     Local verification using configured secret/public key
    /// </summary>
    Local,

    /// <summary>
    ///     Trust API Gateway has verified the token
    /// </summary>
    Gateway,

    /// <summary>
    ///     OAuth 2.0 Token Introspection
    /// </summary>
    Introspection,

    /// <summary>
    ///     Call custom endpoint to verify
    /// </summary>
    Custom
}
