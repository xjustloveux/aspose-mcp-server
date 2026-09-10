using System.Globalization;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Core.Cleanup;

/// <summary>
///     Keeps the cleanup debt queue moving: records what a publish could not delete, retries it on
///     a timer, and reports what an operator has to deal with.
///     <para>
///         Before this, a failed deletion was written to stderr and forgotten. The publish had
///         already succeeded, so the caller saw nothing wrong while a superseded file stayed on
///         disk with no record and no retry — and a restart lost even the line on stderr
///         (R9-F01).
///     </para>
/// </summary>
public sealed class CleanupDebtService : IHostedService, IDisposable
{
    /// <summary>The file the queue is kept in, inside the configured temp directory.</summary>
    private const string QueueFileName = "aspose_cleanup_debts.json";

    /// <summary>Which staging area the next sweep starts with.</summary>
    private const string StagedInputCursorName = "staged-inputs.cursor";

    /// <summary>The flat root, 256 legacy shards, and 65,536 single-copy nonce buckets.</summary>
    private const int StagedInputAreaCount = 65_793;

    /// <summary>How often the queue is swept.</summary>
    private static readonly TimeSpan SweepInterval = TimeSpan.FromMinutes(5);

    /// <summary>
    ///     Output-side journal directories an operator explicitly selected for one-time legacy
    ///     recovery. These are never inferred from <see cref="ServerConfig.AllowedBasePaths" />.
    /// </summary>
    private readonly IReadOnlyList<string> _legacyJournalRoots;

    /// <summary>
    ///     Guards the registration and the timer, which <see cref="StartAsync" />,
    ///     <see cref="StopAsync" /> and <see cref="Dispose" /> all touch and which a faulting host
    ///     can reach from more than one of them.
    /// </summary>
    private readonly object _lifecycle = new();

    private readonly ILogger<CleanupDebtService> _logger;
    private readonly CleanupDebtQueue _queue;

    /// <summary>
    ///     This host's recovery directory and the key its records are signed with.
    /// </summary>
    /// <remarks>
    ///     One per host, held here rather than in a process-wide property. Two hosts in one process
    ///     each have their own temp directory, and a single global answer meant the second to start
    ///     moved the first's journals and signed its queue with a different key (R18-ARCH01).
    /// </remarks>
    private readonly RecoveryContext _recovery;

    /// <summary>
    ///     Where journals are kept and always looked for.
    ///     <para>
    ///         Separate from the output allowlist on purpose. Recovery used to scan only that
    ///         allowlist, so the default open-world configuration — an empty list — scanned
    ///         nothing and no journal was ever found (R13-F01). This location exists whatever the
    ///         allowlist says.
    ///     </para>
    /// </summary>
    /// <summary>
    ///     The roots recovery may move files under, from configuration.
    ///     <para>
    ///         The journal used to carry its own allowlist and recovery read it back, so a file
    ///         dropped in a scanned directory could name the roots it wanted (R13-SEC01). Authority
    ///         comes from here, and a journal that disagrees is refused rather than obeyed.
    ///     </para>
    /// </summary>
    private readonly IReadOnlyList<string> _trustedRoots;

    /// <summary>
    ///     The sink this service installed into <see cref="BoundedFileBatch.RecordDebt" />, so
    ///     shutdown clears its own rather than whatever is there.
    /// </summary>
    private Action<string, string>? _installedSink;

    private Timer? _timer;

    /// <summary>Creates the service and its queue.</summary>
    /// <param name="sessionManager">Supplies the temp directory the queue file lives in.</param>
    /// <param name="serverConfig">Supplies the roots a deletion may happen under.</param>
    /// <param name="logger">The logger the queue reports through.</param>
    public CleanupDebtService(DocumentSessionManager sessionManager, ServerConfig serverConfig,
        ILogger<CleanupDebtService> logger)
    {
        _logger = logger;
        var recoveryDirectory = sessionManager.Config.TempDirectory;

        // The allowlist, and nothing substituted for it when there is none. This used to fall back
        // to the recovery directory, and since journals live there and destinations do not, the
        // default open-world configuration refused every destination it was asked to restore — a
        // crash recovery that recovered nothing, in the only configuration most deployments use
        // (R17-F02). What authorises recovery is the record's signature; these roots are a second
        // bound for a deployment that has chosen one.
        _trustedRoots = serverConfig.AllowedBasePaths;
        _legacyJournalRoots = serverConfig.LegacyPublishJournalRoots;

        // One context for this host, built here and handed to everything that writes or reads a
        // record. It used to be two process-wide statics that whichever host started last rewrote
        // (R18-ARCH01).
        _recovery = RecoveryContext.For(recoveryDirectory);

        _queue = new CleanupDebtQueue(
            Path.Combine(_recovery.Directory, QueueFileName),
            serverConfig.AllowedBasePaths,
            _recovery)
        {
            // A queue that cannot read, write or find room for a debt has stopped doing the one
            // thing it exists for. Without this the seam existed but nothing in production ever
            // assigned it — only the unit tests did, so the reporting §21.11.1 claimed was real
            // only inside those tests (R10-F01).
            OnQueueProblem = problem => _logger.LogWarning(
                "Cleanup debt queue problem: {Problem}", problem)
        };
    }

    /// <summary>Removes input copies a crashed conversion left in this host's staging area.</summary>
    /// <remarks>
    ///     <see cref="ImmutableInputCopy" /> removes its copy when the conversion finishes; a crash
    ///     in between leaves one behind. Its doc promised this sweep before the sweep existed
    ///     (R20-CNV03). Bounded — at most <see cref="EnumerationCeiling" /> entries examined and
    ///     256 files removed per start, none younger than an hour — and confined to the staging
    ///     tree whose contents are this server's own.
    ///     <para>
    ///         Age is not proof of abandonment, and neither was an instant's exclusive open of the
    ///         copy: a conversion holds no handle on its copy between the scan and the loader, and
    ///         a sweep in another host that probed it in that instant deleted a running
    ///         conversion's input (R22-REC02). A copy is deleted only while this sweep holds its
    ///         lease, which the conversion that owns it holds for its whole lifetime.
    ///     </para>
    /// </remarks>
    /// <summary>The most staged-input entries one start examines; a fixture may lower it.</summary>
    internal static int EnumerationCeiling { get; set; } = 65_536;

    /// <summary>Told how many entries a sweep examined, for a fixture that bounds the work.</summary>
    internal static Action<int>? Examined { get; set; }

    /// <summary>
    ///     Opens one staging area's file enumeration; a fixture wraps it to count every yielded
    ///     entry independently of this service's accounting.
    /// </summary>
    internal static Func<string, IEnumerable<string>>? EnumerateStagedInputFiles { get; set; }

    /// <inheritdoc />
    /// <remarks>
    ///     Releases the registration as well as the timer. <see cref="StopAsync" /> may never run
    ///     — a start-up that faults, a host that is disposed without being stopped — and the sink
    ///     stack is static and process-wide, so a registration left on it outlives this service
    ///     and keeps handing debts to a queue nothing is sweeping (R13-F03).
    /// </remarks>
    public void Dispose()
    {
        Release();
    }

    /// <inheritdoc />
    public Task StartAsync(CancellationToken cancellationToken)
    {
        lock (_lifecycle)
        {
            // Starting twice used to push a second registration and forget the first, leaving it
            // on the process-wide stack for good and abandoning a timer that went on firing
            // (R13-F03). One service, one registration; stop it before starting it again.
            if (_installedSink != null)
            {
                _logger.LogDebug("CleanupDebtService is already started; the request was ignored.");
                return Task.CompletedTask;
            }

            return StartCore();
        }
    }

    /// <inheritdoc />
    public Task StopAsync(CancellationToken cancellationToken)
    {
        Release();
        return Task.CompletedTask;
    }

    /// <summary>Registers the sink, recovers, and starts sweeping.</summary>
    /// <returns>A completed task, as the host expects.</returns>
    private Task StartCore()
    {
        // Attached before the first request: a debt incurred and never recorded is the defect.
        // `RecordDebt` is one delegate for the whole process, so a second host would replace this
        // one silently and either host's shutdown would then clear whichever was installed. The
        // replacement is announced rather than silent, and the release below only clears this
        // service's own sink (R10-F01).
        _installedSink = _queue.Record;

        // Registered rather than assigned: the sinks are a stack, so this one is current while it
        // runs and whoever was there before it becomes current again when it stops — in whatever
        // order the hosts happen to shut down (§23.5).
        if (BoundedFileBatch.AddDebtSink(_recovery.Directory, _trustedRoots, _installedSink))
            _logger.LogInformation(
                "Another cleanup debt sink was already registered for this process; debts go to "
                + "this service's queue while it is running, and back to that one afterwards.");

        // The key that separates "a file in this directory" from "a record this server wrote".
        // Established with the context in the constructor; said out loud here, because a host that
        // cannot sign is a host whose crashes will not be rolled back.
        if (_recovery.Capability == null)
            _logger.LogWarning(
                "No recovery signing key could be established in '{Directory}', so publish records "
                + "cannot be written or acted on. A publish interrupted by a crash will leave its "
                + "outputs as they were rather than being rolled back.", _recovery.Directory);

        // Before the sweep, and for the same reason: what a crash left behind is only visible at
        // the next start. A publish that died partway leaves destinations replaced and their
        // previous files beside them; the journal says which, and recovery puts them back
        // (§23.19.2).
        RecoverInterruptedPublishes();
        SweepStagedInputs();

        // Swept once at startup as well as on the timer, because the debts that matter most are
        // the ones that outlived the process that created them.
        Sweep(null);
        _timer = new Timer(Sweep, null, SweepInterval, SweepInterval);

        _logger.LogDebug("CleanupDebtService started; sweeping every {Interval}", SweepInterval);
        return Task.CompletedTask;
    }

    /// <summary>
    ///     Gives up this service's registration and its timer, once, whichever of
    ///     <see cref="StopAsync" /> and <see cref="Dispose" /> gets there first.
    /// </summary>
    private void Release()
    {
        lock (_lifecycle)
        {
            // Only this service's own registration. Clearing the field took away a sink another
            // host had installed, whichever of them stopped first (R10-F01, §23.5). Forgetting it
            // afterwards is what makes a second Stop harmless and a later Start possible.
            if (_installedSink != null)
            {
                BoundedFileBatch.RemoveDebtSink(_recovery.Directory, _installedSink);
                _installedSink = null;
            }

            // Disposed rather than paused. A paused timer that a later Start replaced was never
            // collected, and pausing left nothing for Dispose to distinguish (R13-F03). A sweep
            // already running finishes; it takes no locks this waits on.
            _timer?.Dispose();
            _timer = null;
        }
    }

    private void SweepStagedInputs()
    {
        var staging = Path.Combine(_recovery.Directory, ImmutableInputCopy.DirectoryName);
        if (!Directory.Exists(staging)) return;

        var cutoff = DateTime.UtcNow.AddHours(-1);
        var removed = 0;

        try
        {
            // New copies occupy one claimed bucket each. The flat root and one-level shards remain
            // as legacy areas. The persisted cursor rotates exact bucket probes, while the manual
            // enumerator charges every yielded entry and never calls MoveNext after the work cap.
            var cursorPath = Path.Combine(_recovery.Directory, StagedInputCursorName);
            var startArea = ReadCursor(cursorPath) % StagedInputAreaCount;
            var examined = 0;
            var work = 0;
            var nextArea = startArea;
            var oldest = new PriorityQueue<FileInfo, long>();

            for (var offset = 0;
                 offset < StagedInputAreaCount && work < EnumerationCeiling;
                 offset++)
            {
                var area = (startArea + offset) % StagedInputAreaCount;
                work++;
                var directory = StagedInputArea(staging, area);
                nextArea = (area + 1) % StagedInputAreaCount;
                if (!Directory.Exists(directory)) continue;

                var enumerable = EnumerateStagedInputFiles?.Invoke(directory)
                                 ?? (area >= 257
                                     ? Directory.EnumerateFiles(directory, "input*")
                                     : Directory.EnumerateFiles(directory));
                using var entries = enumerable.GetEnumerator();
                var areaComplete = false;
                while (work < EnumerationCeiling)
                {
                    if (!entries.MoveNext())
                    {
                        areaComplete = true;
                        break;
                    }

                    work++;
                    examined++;
                    var info = new FileInfo(entries.Current);
                    if (info.LinkTarget != null || info.LastWriteTimeUtc > cutoff
                                                || info.Name.EndsWith(ImmutableInputCopy.LeaseSuffix,
                                                    StringComparison.Ordinal))
                        continue;

                    // The newest kept entry sits at the head, so it is the one dropped when the
                    // heap is over the cap: what remains is the oldest 256 examined this start.
                    oldest.Enqueue(info, -info.LastWriteTimeUtc.Ticks);
                    if (oldest.Count > 256) oldest.Dequeue();
                }

                if (!areaComplete)
                {
                    // Do not skip the unobserved tail. New buckets contain only one copy, while a
                    // legacy area is revisited until its old prefix is removed or becomes stale.
                    nextArea = area;
                    break;
                }
            }

            Examined?.Invoke(examined);
            WriteCursor(cursorPath, nextArea);

            var stale = new List<FileInfo>(oldest.Count);
            while (oldest.Count > 0) stale.Add(oldest.Dequeue());
            stale.Sort((a, b) => a.LastWriteTimeUtc.CompareTo(b.LastWriteTimeUtc));

            foreach (var info in stale)
                try
                {
                    // The lease is the proof, held while the copy is deleted: a conversion that
                    // owns this copy holds it and this open fails; a crashed one's lease closed
                    // with its process (R22-REC02). The copy's own exclusive open stays as well —
                    // a copy somebody has open at this instant is not deleted either.
                    using (ImmutableInputCopy.OpenLease(info.FullName, FileMode.OpenOrCreate))
                    {
                        using (new FileStream(info.FullName, FileMode.Open, FileAccess.Read,
                                   FileShare.None))
                        {
                            // A successful exclusive open proves nobody else is using the copy.
                        }

                        info.Delete();
                        removed++;
                    }

                    ImmutableInputCopy.ReleaseBucketFor(info.FullName);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    _logger.LogDebug(ex,
                        "A staged input copy is in use or could not be removed: {File}",
                        info.FullName);
                }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.LogDebug(ex, "The staged-input directory could not be listed");
        }

        if (removed > 0)
            _logger.LogInformation(
                "Removed {Count} input cop(ies) a previous conversion left in '{Directory}'.",
                removed, staging);
    }

    /// <summary>Maps a persisted area number to its legacy directory or exact two-level bucket.</summary>
    /// <param name="staging">The staged-input root.</param>
    /// <param name="area">Zero-based area number.</param>
    /// <returns>The directory represented by the area.</returns>
    private static string StagedInputArea(string staging, int area)
    {
        if (area == 0) return staging;
        if (area <= 256)
            return Path.Combine(staging, (area - 1).ToString("x2",
                CultureInfo.InvariantCulture));

        var bucket = area - 257;
        return Path.Combine(staging,
            (bucket / 256).ToString("x2", CultureInfo.InvariantCulture),
            (bucket % 256).ToString("x2", CultureInfo.InvariantCulture));
    }

    /// <summary>Reads which staging area the next sweep starts with.</summary>
    /// <param name="path">The cursor file.</param>
    /// <returns>The area index, or zero when there is none or it cannot be read.</returns>
    private static int ReadCursor(string path)
    {
        try
        {
            return File.Exists(path) && int.TryParse(File.ReadAllText(path), out var skip) && skip > 0 ? skip : 0;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return 0;
        }
    }

    /// <summary>Writes which staging area the next sweep starts with.</summary>
    /// <param name="path">The cursor file.</param>
    /// <param name="area">The next legacy area or exact two-level nonce bucket.</param>
    private void WriteCursor(string path, int area)
    {
        try
        {
            SecureFile.ReplaceAtomically(path, area.ToString(CultureInfo.InvariantCulture));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.LogDebug(ex, "The staged-input sweep cursor could not be written");
        }
    }

    /// <summary>
    ///     Puts back what a publish interrupted by a crash had already replaced.
    /// </summary>
    private void RecoverInterruptedPublishes()
    {
        // The queue's own directory is always scanned, and it is where journals go. Scanning only
        // the output allowlist meant the default open-world configuration — an empty allowlist —
        // ran this loop zero times and never found a journal at all (R13-F01).
        // Without a key nothing on disk can be told from anything else on disk, so there is no
        // safe reading of the records at all.
        if (_recovery.Capability is not { } capability) return;

        // Journals have been written only under this private directory since recovery was
        // centralised. Allowed output roots are caller-writable namespaces, not recovery
        // authority: scanning them let arbitrary suffix-matching files consume discovery work
        // and starve genuine journals. Older records left in output roots therefore require an
        // explicit operator migration rather than being interpreted automatically.
        var budget = new PublishJournal.RecoveryBudget();

        // One claim for the whole start-up, on this host's own directory, and its wait charged to
        // the same deadline as the scan. Taking it per root wrote a claim file into every legacy
        // allowlisted root — an operator's output directory (R20-REC10).
        using var claim = CrossProcessFileGate.TryAcquire(
            Path.Combine(_recovery.Directory, PublishJournal.ClaimName), budget.Remaining);
        if (claim == null)
        {
            _logger.LogWarning(
                "Another process is recovering '{Directory}', so this start did not.",
                _recovery.Directory);
            return;
        }

        try
        {
            var comparison = OperatingSystem.IsWindows()
                ? StringComparer.OrdinalIgnoreCase
                : StringComparer.Ordinal;
            var roots = new[] { _recovery.Directory }
                .Concat(_legacyJournalRoots)
                .Select(Path.GetFullPath)
                .Distinct(comparison);

            foreach (var root in roots)
            {
                var legacy = !comparison.Equals(root, _recovery.Directory);
                var result = PublishJournal.RecoverAll(root, _trustedRoots, capability,
                    budget, _recovery.Directory);

                if (result.Restored.Count > 0)
                {
                    if (legacy)
                        _logger.LogWarning(
                            "An explicitly configured legacy publish-journal root '{Directory}' "
                            + "restored {Count} file(s) to their previous versions. Remove the "
                            + "legacy option after its journals are drained.",
                            root, result.Restored.Count);
                    else
                        _logger.LogWarning(
                            "A publish did not finish before this server last stopped; {Count} "
                            + "file(s) were restored to the version they had before it started.",
                            result.Restored.Count);
                }

                if (result.Orphaned.Count > 0)
                    _logger.LogWarning(
                        "A publish that did not finish had already created {Count} file(s); they "
                        + "were left in place: {Files}",
                        result.Orphaned.Count, string.Join(", ", result.Orphaned.Take(10)));

                foreach (var refusal in result.Refused)
                    _logger.LogWarning(
                        "An interrupted publish could not be undone and its record was kept for "
                        + "the next start: {Refusal}", refusal);
            }
        }
        catch (Exception ex)
        {
            // Recovery is housekeeping at startup. Failing it must never stop the host from
            // coming up — a server that will not start is worse than one file left as it was.
            _logger.LogError(ex, "Recovering interrupted publishes under {Root} failed",
                _recovery.Directory);
        }
    }

    /// <summary>Runs one sweep and reports what it did.</summary>
    /// <param name="state">Unused timer state.</param>
    private void Sweep(object? state)
    {
        try
        {
            var result = _queue.Sweep();

            if (result.Deleted.Count > 0)
                _logger.LogInformation(
                    "Removed {Count} superseded file(s) a publish could not delete", result.Deleted.Count);

            if (result.Retrying.Count > 0)
                _logger.LogDebug(
                    "{Count} superseded file(s) still could not be removed; will retry",
                    result.Retrying.Count);

            // The one an operator has to act on: it has been retried for the whole retention and
            // is still there.
            foreach (var path in result.Abandoned)
                LogAbandonedCleanup(path);

            foreach (var path in result.Refused)
                LogRefusedCleanup(path);
        }
        catch (Exception ex)
        {
            // A sweep is housekeeping. Failing it must never take the host down.
            _logger.LogError(ex, "Cleanup debt sweep failed");
        }
    }

    private void LogAbandonedCleanup(string path)
    {
        _logger.LogWarning(
            "A superseded file could not be removed within the retention period and is still "
            + "on disk: {Path}. It holds content a newer version replaced.", path);
    }

    private void LogRefusedCleanup(string path)
    {
        _logger.LogWarning(
            "A queued cleanup was refused because the path is no longer one this server may "
            + "delete: {Path}", path);
    }
}
