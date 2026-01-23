namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Result of a recovery operation returned by RecoverSession
/// </summary>
public class RecoverResult
{
    /// <summary>
    ///     Session ID that was recovered
    /// </summary>
    public string SessionId { get; set; } = "";

    /// <summary>
    ///     Whether recovery was successful
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    ///     Path where file was recovered to
    /// </summary>
    public string? RecoveredPath { get; set; }

    /// <summary>
    ///     Original file path before disconnection
    /// </summary>
    public string? OriginalPath { get; set; }

    /// <summary>
    ///     Error message if recovery failed
    /// </summary>
    public string? ErrorMessage { get; set; }
}
