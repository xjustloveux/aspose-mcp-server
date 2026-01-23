namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Authentication result from JWT validation
/// </summary>
public class JwtAuthResult
{
    /// <summary>
    ///     Indicates whether the JWT token is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    ///     Group identifier extracted from the token
    /// </summary>
    public string? GroupId { get; set; }

    /// <summary>
    ///     User identifier extracted from the token (for audit and logging)
    /// </summary>
    public string? UserId { get; set; }

    /// <summary>
    ///     Error message if validation failed
    /// </summary>
    public string? ErrorMessage { get; set; }
}
