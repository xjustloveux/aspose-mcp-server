using System.Globalization;
using AsposeMcpServer.Errors;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Produces a whole request's worth of files somewhere else first, and puts them at their
///     destinations only once every one of them is complete and the request is within budget.
///     <para>
///         <see cref="BoundedFilePublisher" /> makes one file atomic, which is what the fan-out
///         handlers used: each output was published the moment it was produced, so a request that
///         ran out of budget on its last record had already replaced the destinations of every
///         record before it. That is a defensible contract, but it was not the one the tools
///         documented or the one a caller retrying a refused request would expect, and no two of
///         them stated it (R4-R02, R5-R01). A request either produces all of its files or none.
///     </para>
///     <para>
///         Rolling back covers every failure the batch can observe — a write that passes the
///         budget, a serializer that throws, a destination that cannot be replaced. A process that
///         dies between two moves is covered by <see cref="PublishJournal" /> instead: the steps
///         are recorded before they are taken, and the next start puts back whatever an interrupted
///         publish had moved. A handled failure can still leave behind an output directory the
///         handler created before staging began; the tools say so.
///     </para>
///     <para>
///         Two limits are worth stating rather than leaving to be inferred. This survives a process
///         crash, not a power loss: neither the journal nor the staging files are flushed to the
///         device, and no directory is fsynced, so what the operating system had not yet written is
///         gone and the journal may describe steps whose effects did not reach the disk
///         (R18-CONTRACT02). And "failed" does not always mean "nothing changed" — once every
///         output is in place the batch has published, and a record that can then be neither
///         committed nor removed raises <see cref="PublishIndeterminateException" />, which must not
///         be retried (R18-CONTRACT01). Every other failure means the destinations were put back.
///     </para>
/// </summary>
/// <param name="budgetBytes">Bytes the whole request may produce.</param>
/// <param name="what">What is being written, for the error messages.</param>
/// <param name="allowedBasePaths">
///     The allowlist every path this batch touches is re-checked against. A destination validated
///     once says nothing about the sibling <c>.partial-</c> and <c>.replaced-</c> paths derived
///     from it, nor about what its parent directory has become since: each create, move and restore
///     resolves symlinks and re-checks containment immediately before it acts (R5-R01). Empty means
///     no allowlist is configured, which is what the resolver already treats as unrestricted.
/// </param>
/// <param name="recovery">
///     Where this host's publish records live and the key they are signed with. Required rather
///     than read from a process-wide property: two hosts in one process each have their own, and a
///     single global one meant the second to start moved the first's journals and signed with a
///     different key (R18-ARCH01).
/// </param>
public sealed class BoundedFileBatch(
    long budgetBytes,
    string what,
    RecoveryContext recovery,
    IReadOnlyList<string>? allowedBasePaths = null) : IDisposable
{
    private const string StagingPathParameter = "stagingPath";

    /// <summary>
    ///     Debt sinks by the recovery root they belong to.
    ///     <para>
    ///         A debt is a file one host's publish could not remove, to be retried against that
    ///         host's allowlist and recorded in its queue under its key. Reading a process-global
    ///         "most recently registered" sink sent it to whichever host happened to start last
    ///         (R19-REC06).
    ///     </para>
    /// </summary>
    private static readonly Dictionary<string, List<(string Allowlist, Action<string, string> Sink)>>
        SinksByRoot =
            new(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);

    /// <summary>
    ///     The registered debt sinks, most recently installed last.
    ///     <para>
    ///         One delegate for the whole process meant two hosts could not coexist: the second to
    ///         start replaced the first, and whichever stopped first cleared the field — so a host
    ///         that was still running recorded nothing at all. Measured on the current binary: the
    ///         second host stopping left the sink null while the first was still started
    ///         (§23.5). Kept as a stack, so releasing one restores whoever was there before it.
    ///     </para>
    /// </summary>
    private static readonly List<Action<string, string>> Sinks = [];

    /// <summary>Guards <see cref="Sinks" />, which hosts start and stop from any thread.</summary>
    private static readonly object SinkGate = new();

    /// <summary>Paths this batch failed to delete, with the reason for each.</summary>
    private readonly List<(string Path, string Error)> _debts = [];

    // Windows and macOS treat two names differing only in case as one file; Linux does not, and
    // comparing case-insensitively there refused a legitimate `a.txt` alongside `A.txt` (R7-F02).
    private readonly HashSet<string> _destinations = new(
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal);

    private readonly List<(string Staging, string Destination)> _staged = [];
    private bool _published;

    /// <summary>Bytes staged so far across the whole request.</summary>
    public long WrittenBytes { get; private set; }

    /// <summary>How many files are staged.</summary>
    public int Count => _staged.Count;

    /// <summary>
    ///     What could not be put back after a failed publish, if anything. Empty when the rollback
    ///     restored every destination it had already changed.
    /// </summary>
    public IReadOnlyList<string> RollbackFailures { get; private set; } = [];

    /// <summary>
    ///     Backups left behind after a publish that otherwise succeeded. The outputs are all in
    ///     place; these are files the caller may want to remove by hand.
    /// </summary>
    public IReadOnlyList<string> CleanupFailures { get; private set; } = [];

    /// <summary>
    ///     Invoked with (source, destination) immediately before every move this batch makes.
    ///     <para>
    ///         A rollback can only be tested by making one specific move fail — the stage move that
    ///         follows a backup already taken — and no filesystem offers that on demand. Unset in
    ///         production, where it costs a null check (R5-R01).
    ///     </para>
    /// </summary>
    internal Action<string, string>? BeforeMove { get; set; }

    /// <summary>
    ///     Invoked with the path immediately before every delete this batch makes.
    ///     <para>
    ///         Cleanup after the commit point is the one failure a filesystem will not produce on
    ///         demand, and it is the one that used to turn a completed publish into a reported
    ///         failure. Unset in production (R7-F02).
    ///     </para>
    /// </summary>
    internal Action<string>? BeforeDelete { get; set; }

    /// <summary>
    ///     Where an unclaimed cleanup debt is announced. Defaults to the server's diagnostic
    ///     channel; a fixture replaces it to observe the announcement without touching the console.
    /// </summary>
    internal Action<string>? ReportDebt { get; set; }

    /// <summary>
    ///     Where a file that could not be deleted is recorded for later retry, as a path and a
    ///     reason rather than one prose line.
    ///     <para>
    ///         <see cref="CleanupFailures" /> is what a caller is told; this is what a reaper can
    ///         act on. Announcing on stderr was the whole of the follow-up before, so a file left
    ///         behind stayed behind (R9-F01).
    ///     </para>
    /// </summary>
    /// <remarks>
    ///     A fixture's seam, and only that: production never assigns it, and reads it only when no
    ///     host has registered for a batch's root and allowlist. It is not deleted — §33.6 said it
    ///     was, and that was not true — because three fixture files drive the handoff through it;
    ///     what changed is that a live host can no longer be shadowed by it (R20-REC06).
    /// </remarks>
    internal static Action<string, string>? RecordDebt
    {
        get
        {
            lock (SinkGate)
            {
                return Sinks.Count > 0 ? Sinks[^1] : null;
            }
        }
        set
        {
            lock (SinkGate)
            {
                Sinks.Clear();
                if (value != null) Sinks.Add(value);
            }
        }
    }

    /// <summary>
    ///     Removes anything still staged. A published batch has nothing left to remove.
    /// </summary>
    public void Dispose()
    {
        if (_published)
        {
            AnnounceDebts();
            return;
        }

        foreach (var (staging, _) in _staged)
            try
            {
                // Through the seam like every other delete in this class. Calling File.Delete
                // directly here meant a fixture could not reach the one delete that produces a
                // cleanup debt, so the debt path had no test of its own (R9-F01).
                var target = Canonical(staging, StagingPathParameter);
                if (File.Exists(target)) Delete(target);
            }
            catch (Exception error)
                when (error is IOException or UnauthorizedAccessException or ArgumentException)
            {
                // Disposing is the last thing that happens on a failure path; a staging file that
                // cannot be removed is noted rather than thrown, so it does not replace whatever
                // failure brought us here.
                CleanupFailures = CleanupFailures
                    .Append($"'{Path.GetFileName(staging)}' was left behind: {error.Message}")
                    .ToList();
                _debts.Add((staging, error.Message));
            }

        _staged.Clear();
        _destinations.Clear();
        AnnounceDebts();
    }

    /// <summary>The identity of an allowlist, so two hosts on one root can be told apart.</summary>
    /// <param name="allowedBasePaths">The allowlist.</param>
    /// <returns>A canonical, order-independent key; empty for an unrestricted host.</returns>
    /// <remarks>
    ///     Two hosts sharing a recovery root with different allowlists were the same host to the
    ///     registry, so a debt could be handed to the one whose allowlist did not cover it
    ///     (R20-REC06). The allowlist is what decides which host may delete, so it is part of the
    ///     key.
    /// </remarks>
    internal static string SinkKey(IReadOnlyList<string> allowedBasePaths)
    {
        var comparer = OperatingSystem.IsWindows()
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal;

        // Length-prefixed, so nothing a path contains can read as the boundary between two of
        // them: joined on a newline alone, the one Unix path "/a\n/b" and the two paths "/a",
        // "/b" made the same key and two hosts one sink (R22-REC03).
        return string.Join("\n", allowedBasePaths
            .Select(path =>
            {
                try
                {
                    var canonical = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
                    // Canonical casing where the filesystem ignores it: sorting case-insensitively
                    // and then comparing the joined string with `==` made `C:\\X` and `c:\\x` two
                    // hosts (R21-REC07).
                    return OperatingSystem.IsWindows() ? canonical.ToUpperInvariant() : canonical;
                }
                catch (Exception ex) when (ex is ArgumentException or NotSupportedException
                                               or PathTooLongException or IOException)
                {
                    return path;
                }
            })
            .OrderBy(path => path, comparer)
            .Select(path => path.Length.ToString(CultureInfo.InvariantCulture) + ":" + path));
    }

    /// <summary>Registers a sink for one recovery root and allowlist.</summary>
    /// <param name="recoveryDirectory">The root whose debts this sink records.</param>
    /// <param name="allowedBasePaths">The allowlist the registering host deletes under.</param>
    /// <param name="sink">The sink.</param>
    /// <returns><c>true</c> when another host had already registered for that root and allowlist.</returns>
    internal static bool AddDebtSink(string recoveryDirectory, IReadOnlyList<string> allowedBasePaths,
        Action<string, string> sink)
    {
        var key = SinkKey(allowedBasePaths);

        lock (SinkGate)
        {
            if (!SinksByRoot.TryGetValue(recoveryDirectory, out var sinks))
                SinksByRoot[recoveryDirectory] = sinks = [];

            var displaced = sinks.Any(entry => entry.Allowlist == key);
            sinks.Add((key, sink));
            return displaced;
        }
    }

    /// <summary>Registers a sink for an unrestricted host on one recovery root.</summary>
    /// <param name="recoveryDirectory">The root whose debts this sink records.</param>
    /// <param name="sink">The sink.</param>
    /// <returns><c>true</c> when another unrestricted host had already registered for that root.</returns>
    internal static bool AddDebtSink(string recoveryDirectory, Action<string, string> sink)
    {
        return AddDebtSink(recoveryDirectory, [], sink);
    }

    /// <summary>Releases a sink, leaving any other host's registration alone.</summary>
    /// <param name="recoveryDirectory">The root it was registered for.</param>
    /// <param name="sink">The sink to release.</param>
    internal static void RemoveDebtSink(string recoveryDirectory, Action<string, string> sink)
    {
        lock (SinkGate)
        {
            if (!SinksByRoot.TryGetValue(recoveryDirectory, out var sinks)) return;

            sinks.RemoveAll(entry => ReferenceEquals(entry.Sink, sink));
            if (sinks.Count == 0) SinksByRoot.Remove(recoveryDirectory);
        }
    }

    /// <summary>The sink that owns a recovery root's debts, if a host has registered one.</summary>
    /// <param name="recoveryDirectory">The root the debt belongs to.</param>
    /// <returns>The sink, or null when no host owns that root.</returns>
    /// <remarks>
    ///     Null is the honest answer for a batch whose host is not running, and it falls back to
    ///     the announcement rather than to another host's queue: recording a debt somewhere it
    ///     will be swept under a different allowlist is worse than not recording it.
    /// </remarks>
    internal static Action<string, string>? DebtSinkFor(string recoveryDirectory)
    {
        lock (SinkGate)
        {
            return SinksByRoot.TryGetValue(recoveryDirectory, out var sinks) && sinks.Count > 0
                ? sinks[^1].Sink
                : null;
        }
    }

    /// <summary>The sink of the host that owns a root <em>and</em> deletes under this allowlist.</summary>
    /// <param name="recoveryDirectory">The root the debt belongs to.</param>
    /// <param name="allowedBasePaths">The allowlist the batch that incurred the debt ran under.</param>
    /// <returns>The matching host's sink, or null when no such host is registered.</returns>
    /// <remarks>
    ///     Exact match on the allowlist, most recently registered first. A host with a different
    ///     allowlist is not "close enough": handing it a debt it may not delete only gets the debt
    ///     refused there, and until R20-REC06 refused meant dropped.
    /// </remarks>
    internal static Action<string, string>? DebtSinkFor(string recoveryDirectory,
        IReadOnlyList<string> allowedBasePaths)
    {
        var key = SinkKey(allowedBasePaths);

        lock (SinkGate)
        {
            if (!SinksByRoot.TryGetValue(recoveryDirectory, out var sinks)) return null;

            for (var i = sinks.Count - 1; i >= 0; i--)
                if (sinks[i].Allowlist == key)
                    return sinks[i].Sink;

            return null;
        }
    }

    /// <summary>
    ///     Announces every cleanup debt this batch is still carrying.
    ///     <para>
    ///         A backup this batch could not delete after committing may hold the previous
    ///         version of the caller's document. Every production caller ignored
    ///         <see cref="CleanupFailures" />, and the single-file wrapper dropped the batch
    ///         entirely, so such a file stayed on disk with nothing said about it anywhere
    ///         (R8-F02). Saying so costs nothing and is what makes it possible to clean up later.
    ///     </para>
    /// </summary>
    private void AnnounceDebts()
    {
        foreach (var debt in CleanupFailures)
        {
            var message = $"[WARN] {what}: a file this request replaced could not be removed and "
                          + $"is still on disk: {debt}";

            if (ReportDebt != null) ReportDebt(message);
            else Console.Error.WriteLine(message);
        }

        // Handed on for retry as well as announced. Announcing alone made the debt visible once
        // and never acted on it (R9-F01).
        // The sink of the host that owns this batch's root *and* deletes under this batch's
        // allowlist. Root alone made two differently-allowlisted hosts interchangeable
        // (R20-REC06). The process-wide seam is only consulted when no host matches at all.
        var record = DebtSinkFor(recovery.Directory, allowedBasePaths ?? []) ?? RecordDebt;
        if (record == null) return;

        foreach (var (path, error) in _debts) record(path, error);
    }

    /// <summary>
    ///     Deletes a file, giving a fixture the chance to fail this exact delete.
    /// </summary>
    /// <param name="path">What to delete.</param>
    private void Delete(string path)
    {
        BeforeDelete?.Invoke(path);
        File.Delete(path);
    }

    /// <summary>
    ///     Moves a file, giving a fixture the chance to fail this exact move.
    /// </summary>
    /// <param name="source">What to move.</param>
    /// <param name="destination">Where to move it.</param>
    /// <param name="overwrite">Whether an existing destination may be replaced.</param>
    private void Move(string source, string destination, bool overwrite)
    {
        BeforeMove?.Invoke(source, destination);
        File.Move(source, destination, overwrite);
    }

    /// <summary>
    ///     Re-resolves a path and confirms it is still inside the allowlist.
    /// </summary>
    /// <param name="path">The path about to be created, moved or restored.</param>
    /// <param name="paramName">What the path is, for the refusal message.</param>
    /// <returns>The canonical path to act on.</returns>
    /// <exception cref="ArgumentException">Thrown when the path is outside the allowlist.</exception>
    private string Canonical(string path, string paramName)
    {
        return SecurityHelper.ResolveAndEnsureWithinAllowlist(path, allowedBasePaths ?? [], paramName);
    }

    /// <summary>
    ///     Produces one file of the request, without touching its destination.
    /// </summary>
    /// <param name="destination">Final path, already resolved against the allowlist.</param>
    /// <param name="write">Writes the content to the stream it is given.</param>
    /// <returns>The number of bytes this file holds.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the file passes what the request has left, when the same destination is
    ///     staged twice, or when a path this batch must touch is not inside the allowlist. Nothing
    ///     is published, and <see cref="Dispose" /> removes what was staged before it.
    /// </exception>
    public long Stage(string destination, Action<Stream> write)
    {
        // A batch is one transaction. Staging into a published one produced a file nothing would
        // ever publish and Dispose would not clean up, because the batch counted as finished
        // (R7-F02).
        if (_published)
            throw new ArgumentException(
                $"This request has already published its {what}; it cannot produce more.");

        var remaining = budgetBytes - WrittenBytes;
        if (remaining <= 0)
            throw new ArgumentException(
                $"The request has already produced the maximum of {what} it may, so '"
                + Path.GetFileName(destination) + "' was not written. Extract a subset instead.");

        // Reservation, staging-name derivation and the write are one unit. Deriving the staging
        // name outside the try meant a refusal there left the destination reserved for a file this
        // batch would never produce, and the batch could not be retried (R8-F01).
        var canonicalDestination = Canonical(destination, nameof(destination));
        string? reserved = null;
        string? staging = null;

        var completed = false;
        try
        {
            // Two entries for one destination would make the publish order decide the winner and
            // the rollback restore the wrong file.
            if (!_destinations.Add(canonicalDestination))
                throw new ArgumentException(
                    $"'{Path.GetFileName(canonicalDestination)}' is already being written by this "
                    + "request, so one of the two would silently replace the other.");

            reserved = canonicalDestination;

            // A full identifier, not a prefix: a staging name that can be guessed is a name that
            // can be planted in advance.
            staging = Canonical(canonicalDestination + ".partial-" + Guid.NewGuid().ToString("N"),
                StagingPathParameter);

            // CreateNew rather than Create: if anything already sits at this name, that is not a
            // file this batch made and it must not be overwritten.
            // ReadWrite, not Write: several document libraries seek back over what they have
            // written while saving, and a write-only stream fails them with "Stream was not
            // readable" (R7-C01). The bound still applies to every write.
            using (var file = new FileStream(staging, FileMode.CreateNew, FileAccess.ReadWrite))
            using (var bounded = new BoundedWriteStream(file, remaining, what))
            {
                try
                {
                    write(bounded);
                }
                catch (Exception ex) when (ex is not ArgumentException && bounded.Refused)
                {
                    // The bound was hit inside the writer, and the writer reported it in its
                    // own words — Aspose turns a stream's refusal into "Not supported image
                    // type". The refusal is the budget's, and is reported as such (R23-PDF01).
                    throw new ArgumentException(
                        $"The request has produced more than the {budgetBytes:N0} bytes of {what} it may. "
                        + "Extract a subset instead.", ex);
                }
            }

            var length = new FileInfo(staging).Length;
            WrittenBytes += length;
            _staged.Add((staging, canonicalDestination));
            completed = true;
            return length;
        }
        finally
        {
            if (!completed)
            {
                if (reserved != null) _destinations.Remove(reserved);
                CleanUpStaging(staging);
            }
        }
    }

    /// <summary>
    ///     Removes a staging file a failed <see cref="Stage" /> left behind.
    ///     <para>
    ///         The delete went straight to <see cref="File.Delete(string)" />: no re-resolution, no
    ///         seam for a fixture to fail it, and any error from it replaced the failure that
    ///         brought us here. It now goes through the same canonicalisation as every other path
    ///         this class touches, and a cleanup that cannot be done is recorded rather than
    ///         raised, so the original failure is still the one the caller sees (R8-F01).
    ///     </para>
    /// </summary>
    /// <param name="staging">The staging path, or <c>null</c> when there is not one yet.</param>
    private void CleanUpStaging(string? staging)
    {
        if (staging == null) return;

        try
        {
            var canonical = Canonical(staging, StagingPathParameter);
            if (File.Exists(canonical)) Delete(canonical);
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException
                                          or ArgumentException)
        {
            CleanupFailures = CleanupFailures.Append(
                $"{staging}: {error.Message}").ToList();
            _debts.Add((staging, error.Message));
        }
    }

    /// <summary>
    ///     Moves every staged file to its destination, or leaves every destination as it was.
    /// </summary>
    /// <returns>The destinations written, in the order they were staged.</returns>
    /// <exception cref="AggregateException">
    ///     Thrown when a destination could not be replaced <em>and</em> the rollback could not put
    ///     back what it had already changed. It carries the original failure first, then each
    ///     rollback failure, because a destination that is now missing is not the same event as a
    ///     refusal that left everything alone.
    /// </exception>
    public IReadOnlyList<string> Publish()
    {
        // Every step is written down before it is taken. Recording the pair only after both moves
        // succeeded meant the one in flight was not in the journal, so a stage move that failed
        // after its backup had been taken left the destination missing (R5-R01).
        var journal = new List<PublishStep>();

        // The in-process rollback below handles a move that fails. It cannot handle the process
        // dying: nothing runs, and what is left is some destinations replaced, their previous
        // files sitting as `.replaced-*` siblings, and no record of which was which. The journal
        // is that record — written before the first move, removed after the last, so its presence
        // at startup means a publish did not finish (§23.19.2).
        var journalPath = JournalPath();

        try
        {
            // The whole plan is written before the first move. Writing it per step meant a
            // batch that only created files never wrote one at all, and a crash before the
            // second step had no record of the first (R13-F01).
            foreach (var (staging, destination) in _staged)
            {
                var step = new PublishStep(Canonical(destination, "destination"));

                if (File.Exists(step.Destination))
                    step.Backup = Canonical(
                        step.Destination + ".replaced-" + Guid.NewGuid().ToString("N"),
                        "backupPath");
                else
                    // A file this publish creates. Recording what it will contain is what lets
                    // recovery remove it after a crash without becoming a way to delete any file
                    // a journal cares to name; the staged copy already holds the final bytes.
                    step.Created = PublishJournal.DigestOf(staging);

                journal.Add(step);
            }

            WriteJournal(journalPath, journal, PublishJournal.Publishing);

            for (var i = 0; i < _staged.Count; i++)
            {
                var step = journal[i];
                var staging = _staged[i].Staging;

                if (step.Backup != null) Move(step.Destination, step.Backup, false);

                Move(Canonical(staging, StagingPathParameter), step.Destination, true);
                step.Placed = true;
            }
        }
        catch (Exception failure)
        {
            var rollbackFailures = RollBack(journal);
            RollbackFailures = rollbackFailures;

            // The rollback ran, so there is nothing for a later recovery to put back. A journal
            // left here would make the next start restore backups that no longer exist.
            if (rollbackFailures.Count == 0) PublishJournal.Delete(journalPath);
            if (rollbackFailures.Count == 0) throw;

            throw new AggregateException(
                $"Publishing {what} failed, and the destinations could not all be put back.",
                new[] { failure }.Concat(
                    rollbackFailures.Select(message => (Exception)new IOException(message))));
        }

        // Every output is in place: this is the commit point. Removing the backups is tidying up
        // after a request that has already succeeded, so a failure here is reported as a warning
        // rather than turning a completed publish into a failure the caller would retry (R7-F02).
        //
        // But the record has to say so before anything can be reported as succeeded. It used to be
        // deleted here, with the failure swallowed, so a publish whose journal could not be removed
        // left something the next start read as interrupted — and rolled back, deleting the very
        // files it had just delivered (R17-F01). Marked committed first, then deleted; if neither
        // can be made to stick, the caller is told, because at that point nothing can stop a later
        // start from undoing the work.
        // Set before the record is settled, not after. Every output is in place, so this batch
        // has published whatever happens next; leaving it false meant an indeterminate commit came
        // back to a caller whose batch would happily publish again (R18-CONTRACT01).
        _published = true;

        Commit(journalPath, journal);

        var cleanupFailures = new List<string>();
        foreach (var step in journal)
        {
            if (step.Backup == null) continue;

            try
            {
                var backup = Canonical(step.Backup, "backupPath");
                if (File.Exists(backup)) Delete(backup);
            }
            catch (Exception error)
                when (error is IOException or UnauthorizedAccessException or ArgumentException)
            {
                cleanupFailures.Add(
                    $"the previous '{Path.GetFileName(step.Destination)}' was kept at "
                    + $"'{Path.GetFileName(step.Backup)}': {error.Message}");

                // The message above names the file for a person; this hands the *path* to the
                // reaper. Only the message was produced before, so the one debt this whole
                // mechanism exists for — the previous version of the caller's document, left on
                // disk after a publish that succeeded — was the one debt never queued for retry
                // (R9-F01).
                _debts.Add((step.Backup, error.Message));
            }
        }

        CleanupFailures = cleanupFailures;
        return journal.Select(step => step.Destination).ToList();
    }

    /// <summary>
    ///     Undoes the steps this publish had taken, most recent first.
    /// </summary>
    /// <param name="journal">What the publish recorded before each step.</param>
    /// <returns>A message for each destination that could not be put back.</returns>
    private List<string> RollBack(List<PublishStep> journal)
    {
        var failures = new List<string>();

        for (var i = journal.Count - 1; i >= 0; i--)
        {
            var step = journal[i];
            try
            {
                // Re-checked here as well: the class promises containment before every create,
                // move and restore, and the rollback used to be the one path that skipped it, so a
                // parent swapped during the failure was followed on the way back (R7-F02).
                var destination = Canonical(step.Destination, "destination");

                // Only a destination this publish actually wrote is removed; one it never reached
                // still holds the caller's own file.
                if (step.Placed && File.Exists(destination)) File.Delete(destination);

                if (step.Backup != null)
                {
                    var backup = Canonical(step.Backup, "backupPath");
                    if (File.Exists(backup)) Move(backup, destination, true);
                }
            }
            catch (Exception error)
                when (error is IOException or UnauthorizedAccessException or ArgumentException)
            {
                failures.Add($"'{Path.GetFileName(step.Destination)}' could not be restored: "
                             + error.Message);
            }
        }

        return failures;
    }

    /// <summary>Where this batch records an in-flight publish.</summary>
    /// <returns>
    ///     The journal path, inside the host's recovery directory — not beside the destinations,
    ///     which is where it lived before §26.2 and what this said until R20-DOC01.
    /// </returns>
    private string? JournalPath()
    {
        // A batch with nothing staged replaces nothing, so there is nothing a crash could leave
        // half-done. Reading _staged[0] unconditionally threw on exactly that case — an
        // extraction that found no images publishes an empty batch.
        if (_staged.Count == 0) return null;

        // Where recovery looks, and nowhere else. This used to be the first allowed root, or —
        // with no allowlist, which is the default — the directory the caller named for its output.
        // Recovery only ever scanned the server's temp directory and the allowlist, so in that
        // default configuration a real crash left a real journal somewhere nothing would look at
        // again (independent review of R13-F01). It then became one process-wide property, which
        // two hosts in one process overwrote for each other (R18-ARCH01); now each batch is handed
        // the context it belongs to.
        var directory = recovery.Directory;

        return Path.Combine(directory, Guid.NewGuid().ToString("N") + PublishJournal.Extension);
    }

    /// <summary>
    ///     Records the steps taken so far, so a crash leaves something recoverable.
    /// </summary>
    /// <param name="path">The journal path.</param>
    /// <param name="journal">The steps taken so far.</param>
    private void Commit(string? path, List<PublishStep> journal)
    {
        if (path == null) return;

        var committed = false;
        try
        {
            WriteJournal(path, journal, PublishJournal.Committed);
            committed = true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Falls through to the delete, which is the other way to reach the same state.
            ReportDebt?.Invoke(
                $"{path}: the publish record could not be marked committed ({ex.Message})");
        }

        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            if (committed)
            {
                // A committed record left behind is harmless: recovery reads it, tidies the
                // backups and removes it, and never touches a destination.
                ReportDebt?.Invoke(
                    $"{path}: the committed publish record could not be removed ({ex.Message})");
                return;
            }

            throw new PublishIndeterminateException(
                $"{what}: every output is in place, but the publish record could neither be marked "
                + "committed nor removed, so a later start would treat this finished publish as an "
                + "interrupted one and undo it. The outputs have been left as published, so this "
                + $"must not be retried; the record at '{path}' needs an operator. ({ex.Message})",
                path, ex);
        }
    }

    /// <summary>Writes the journal in a given state.</summary>
    /// <param name="path">Where the journal lives, or null when this batch writes none.</param>
    /// <param name="journal">The steps taken so far.</param>
    /// <param name="state">Whether the transaction is still publishing or has committed.</param>
    private void WriteJournal(string? path, List<PublishStep> journal, string state)
    {
        if (path == null) return;

        try
        {
            // Every step, not only the ones replacing something. Filtering to steps with a
            // backup meant a batch creating new files journalled nothing at all, so a crash
            // partway through left those files behind with no record (R13-F01).
            // Stamped with this process, so a start-up recovery can tell a crashed publish from
            // one that is still running in another host sharing this directory.
            var capability = recovery.Capability
                             ?? throw new IOException(
                                 "this installation has no recovery signing key, so a publish record could not "
                                 + "be written in a form recovery would act on");

            PublishJournal.WriteManaged(path, PublishJournal.Signed(PublishJournal.Describe(
                what,
                journal.Select(step =>
                        new PublishJournal.Entry(step.Destination, step.Backup, step.Created))
                    .ToList(),
                state), capability));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Refused, not swallowed. Continuing meant the publish went ahead claiming a crash
            // guarantee it no longer had — silently downgrading the property the journal exists
            // for. Nothing has been moved at this point, so failing here leaves the tree
            // untouched (R13-F01).
            throw new IOException(
                $"{what}: the publish journal could not be written, so a crash partway through "
                + "could not be undone. Nothing was published. " + ex.Message, ex);
        }
    }

    /// <summary>
    ///     One destination's place in the publish, written down before the moves that change it.
    /// </summary>
    /// <param name="destination">The canonical destination this step writes.</param>
    private sealed class PublishStep(string destination)
    {
        /// <summary>The canonical destination this step writes.</summary>
        public string Destination { get; } = destination;

        /// <summary>Where the previous file at that destination was moved, if there was one.</summary>
        public string? Backup { get; set; }

        /// <summary>Whether the staged file reached the destination.</summary>
        public bool Placed { get; set; }

        /// <summary>The digest of the bytes this step writes, when it creates the destination.</summary>
        public string? Created { get; set; }
    }
}
