namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Statistics about temp files returned by GetStats
/// </summary>
public class TempFileStats
{
    /// <summary>
    ///     Total number of temp files
    /// </summary>
    public int TotalCount { get; set; }

    /// <summary>
    ///     Total size of all temp files in bytes
    /// </summary>
    public long TotalSizeBytes { get; set; }

    /// <summary>
    ///     Number of expired temp files (past retention period)
    /// </summary>
    public int ExpiredCount { get; set; }

    /// <summary>
    ///     Total size of all temp files in megabytes
    /// </summary>
    public double TotalSizeMb => TotalSizeBytes / (1024.0 * 1024.0);
}
