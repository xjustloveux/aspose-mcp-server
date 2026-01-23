namespace AsposeMcpServer.Core.Security;

/// <summary>
///     API Key authentication verification mode
/// </summary>
public enum ApiKeyMode
{
    /// <summary>
    ///     Local verification using configured keys
    /// </summary>
    Local,

    /// <summary>
    ///     Trust API Gateway has verified the request
    /// </summary>
    Gateway,

    /// <summary>
    ///     Call external API to verify the key
    /// </summary>
    Introspection,

    /// <summary>
    ///     Call custom endpoint to verify
    /// </summary>
    Custom
}
