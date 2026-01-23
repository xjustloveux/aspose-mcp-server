namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Static class containing all document session result types for schema generation.
/// </summary>
public static class DocumentSessionResults
{
    /// <summary>
    ///     All result types used by DocumentSessionTool.
    /// </summary>
    public static readonly Type[] AllTypes =
    [
        typeof(OpenSessionResult),
        typeof(SaveSessionResult),
        typeof(CloseSessionResult),
        typeof(ListSessionsResult),
        typeof(SessionStatusResult),
        typeof(ListTempFilesResult),
        typeof(RecoverTempFileResult),
        typeof(DeleteTempFileResult),
        typeof(CleanupTempFilesResult),
        typeof(TempFileStatsResult)
    ];
}
