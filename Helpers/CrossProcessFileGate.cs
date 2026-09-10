namespace AsposeMcpServer.Helpers;

/// <summary>
///     Mutual exclusion between processes, for a record more than one of them can reach.
///     <para>
///         <see cref="CleanupDebtQueue" /> synchronised its read-modify-write with a
///         <c>lock</c> on a process-local object, which orders threads and nothing else. Two
///         servers pointed at one temp directory share one queue file — that is not a
///         misconfiguration but the point of the queue, since debts have to outlive the process
///         that recorded them — and both read the same list, each added its own debt, and whichever
///         wrote second erased the other's (R18-ARCH02).
///     </para>
///     <para>
///         A lock file opened with <see cref="FileShare.None" />, which .NET carries to a mandatory
///         lock on Windows and an advisory <c>flock</c> on Unix. Held for the whole read-modify-write
///         rather than either half of it, because it is the gap between the read and the write that
///         loses the update.
///     </para>
/// </summary>
public static class CrossProcessFileGate
{
    /// <summary>How long to keep trying before giving up on a gate someone else is holding.</summary>
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(10);

    /// <summary>How long to wait between attempts.</summary>
    private static readonly TimeSpan Retry = TimeSpan.FromMilliseconds(15);

    /// <summary>Takes the gate, waiting for whoever holds it.</summary>
    /// <param name="lockFile">The file whose handle is the gate.</param>
    /// <param name="timeout">How long to wait, or <see cref="DefaultTimeout" /> when omitted.</param>
    /// <returns>A handle to release, or <c>null</c> when the gate could not be taken in time.</returns>
    /// <remarks>
    ///     Null is a refusal to proceed, not permission to proceed unguarded: a caller that could
    ///     not take the gate has to leave the record alone, because the whole reason for the gate
    ///     is that somebody else is part-way through changing it.
    ///     <para>
    ///         Not re-entrant. A second acquire of the same gate on the same thread waits for the
    ///         first to be released and then times out, so a gated operation must not call another.
    ///     </para>
    /// </remarks>
    public static IDisposable? TryAcquire(string lockFile, TimeSpan? timeout = null)
    {
        var deadline = DateTime.UtcNow + (timeout ?? DefaultTimeout);

        var directory = Path.GetDirectoryName(lockFile);
        try
        {
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return null;
        }

        while (true)
            try
            {
                return new FileStream(lockFile, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                    FileShare.None, 1, FileOptions.WriteThrough);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Held by somebody, or not openable at all. Both are answered the same way: wait,
                // then stop waiting. Telling the two apart would need the error code and would not
                // change what happens next.
                if (DateTime.UtcNow >= deadline) return null;
                Thread.Sleep(Retry);
            }
    }
}
