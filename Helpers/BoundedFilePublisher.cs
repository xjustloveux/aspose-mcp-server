namespace AsposeMcpServer.Helpers;

/// <summary>
///     Produces one file somewhere else first, and puts it at the caller's destination only once it
///     is complete and within budget.
///     <para>
///         The extraction paths saved each part straight onto the destination and measured it
///         afterwards, so the file that passed the limit was already on disk when the request was
///         refused, and the caller was left with a partial set of outputs and no way to tell which
///         ones were meant to be there (R3-R07).
///     </para>
///     <para>
///         This is a single-file <see cref="BoundedFileBatch" /> and nothing else. It used to be a
///         parallel implementation with a short staging suffix, no allowlist and no rollback
///         journal, so every hardening the batch received — full identifiers, a destination that
///         cannot be replaced by a name that was guessed, re-checking containment before each
///         create, move and restore — stopped at the six fan-out callers and left the single-file
///         ones behind (R7-F01).
///     </para>
/// </summary>
public static class BoundedFilePublisher
{
    /// <summary>
    ///     Produces a file within a byte budget and publishes it atomically.
    /// </summary>
    /// <param name="destination">Final path, already resolved against the allowlist.</param>
    /// <param name="remainingBytes">
    ///     Bytes still available in the request's budget. A value of zero or less refuses before
    ///     anything is written.
    /// </param>
    /// <param name="write">Writes the content to the stream it is given.</param>
    /// <param name="what">What is being written, for the error message.</param>
    /// <param name="allowedBasePaths">
    ///     The allowlist the destination and every path derived from it are re-checked against, at
    ///     the moment each one is acted on. Empty means no allowlist is configured.
    /// </param>
    /// <returns>What was published, and anything the publish could not clean up afterwards.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the content passes the remaining budget, or when a path this publish must
    ///     touch is not inside the allowlist. The destination is untouched.
    /// </exception>
    /// <param name="recovery">
    ///     Where this host's publish records live and the key they are signed with (R18-ARCH01).
    /// </param>
    public static PublishOutcome Publish(string destination, long remainingBytes, Action<Stream> write,
        string what, RecoveryContext recovery, IReadOnlyList<string> allowedBasePaths)
    {
        using var batch = new BoundedFileBatch(remainingBytes, what, recovery, allowedBasePaths);
        var written = batch.Stage(destination, write);
        batch.Publish();

        return new PublishOutcome(written, batch.CleanupFailures);
    }
}

/// <summary>
///     What one publish produced, and what it left behind.
///     <para>
///         The wrapper used to return a byte count and drop the batch, which threw away the one
///         record of a backup that could not be deleted — a file that may hold the previous version
///         of the caller's document (R8-F02). The count alone could not carry that, so the result
///         says both things.
///     </para>
/// </summary>
/// <param name="WrittenBytes">How many bytes were written.</param>
/// <param name="CleanupFailures">
///     Files this publish replaced but could not remove afterwards. Empty on an ordinary publish;
///     an entry means something is still on disk that nothing else will tidy up.
/// </param>
public readonly record struct PublishOutcome(
    long WrittenBytes,
    IReadOnlyList<string> CleanupFailures);
