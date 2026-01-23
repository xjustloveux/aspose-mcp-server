namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Information about a recoverable file returned by ListRecoverableFiles
/// </summary>
public class RecoverableFileInfo
{
    /// <summary>
    ///     Session ID that created this temp file
    /// </summary>
    public string SessionId { get; set; } = "";

    /// <summary>
    ///     Original file path before disconnection
    /// </summary>
    public string OriginalPath { get; set; } = "";

    /// <summary>
    ///     Path to the temporary file
    /// </summary>
    public string TempPath { get; set; } = "";

    /// <summary>
    ///     Document type (Word, Excel, PowerPoint, Pdf)
    /// </summary>
    public string DocumentType { get; set; } = "";

    /// <summary>
    ///     When the temp file was saved
    /// </summary>
    public DateTime SavedAt { get; set; }

    /// <summary>
    ///     When the temp file will expire and be cleaned up
    /// </summary>
    public DateTime ExpiresAt { get; set; }

    /// <summary>
    ///     Size of the temp file in bytes
    /// </summary>
    public long FileSizeBytes { get; set; }

    /// <summary>
    ///     Whether to prompt user on reconnect
    /// </summary>
    public bool PromptOnReconnect { get; set; }

    /// <summary>
    ///     Owner group ID for session isolation
    /// </summary>
    public string? OwnerGroupId { get; set; }

    /// <summary>
    ///     Owner user ID for audit and logging
    /// </summary>
    public string? OwnerUserId { get; set; }
}
