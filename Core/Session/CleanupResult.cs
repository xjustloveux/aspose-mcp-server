namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Result of a cleanup operation returned by CleanupExpiredFiles
/// </summary>
public class CleanupResult
{
    /// <summary>
    ///     Number of temp files scanned
    /// </summary>
    public int ScannedCount { get; set; }

    /// <summary>
    ///     Number of temp files deleted
    /// </summary>
    public int DeletedCount { get; set; }

    /// <summary>
    ///     Number of errors during cleanup
    /// </summary>
    public int ErrorCount { get; set; }
}
