namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Authentication result from API Key validation
/// </summary>
public class ApiKeyAuthResult
{
    /// <summary>
    ///     Indicates whether the API key is valid
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    ///     Group identifier extracted from the API key
    /// </summary>
    public string? GroupId { get; set; }

    /// <summary>
    ///     Error message if validation failed
    /// </summary>
    public string? ErrorMessage { get; set; }
}
