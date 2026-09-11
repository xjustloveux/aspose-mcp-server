using System.Security;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Remembers files a publish replaced but could not delete, and keeps trying.
///     <para>
///         A transactional publish moves the new file into place and then removes the old one. When
///         that removal fails — the file is open, a scanner has it locked, the volume is
///         momentarily read-only — the publish has already succeeded and the old file stays on
///         disk. It was announced on stderr and forgotten, so a superseded document could remain
///         readable indefinitely with nothing tracking it (R9-F01).
///     </para>
///     <para>
///         This queue survives a restart, retries with backoff, and gives up loudly rather than
///         silently: a debt that outlives its retention is reported as abandoned so an operator can
///         act on it, instead of being retried forever or dropped.
///     </para>
///     <para>
///         Every deletion is re-validated at the moment it happens — canonicalised again, checked
///         against the allowlist again, and refused if the path has become a reparse point. The
///         queue is a record of intent, not an authority: a path in it is a path to <em>consider</em>
///         deleting, and if the filesystem has changed underneath it the answer is to refuse and
///         record why. Nothing is ever deleted because a file said so.
///     </para>
/// </summary>
public sealed class CleanupDebtQueue
{
    /// <summary>Most debts held at once, so a failing volume cannot grow this without bound.</summary>
    public const int MaxDebts = 10_000;

    /// <summary>How many dead entries the registry tolerates before it sweeps them.</summary>
    /// <remarks>
    ///     Swept on a threshold rather than on every construction, so the common case — a server
    ///     with one queue — never walks the dictionary.
    /// </remarks>
    private const int GateSweepThreshold = 64;

    /// <summary>The most a debt's target may be to have its content digested.</summary>
    /// <remarks>
    ///     A debt is a file a publish had just written and could not remove, so it is bounded by
    ///     that publish's own budget in practice. The cap is what stops a sweep reading an
    ///     arbitrarily large file on a queue somebody else filled in.
    /// </remarks>
    public const long MaxDigestBytes = 64L * 1024 * 1024;

    /// <summary>Largest queue file this will read, in bytes.</summary>
    /// <remarks>
    ///     Checked before the file is opened for content. <see cref="MaxDebts" /> bounded what this
    ///     process would <em>record</em> and nothing bounded what it would read back, so the size
    ///     of the work at start-up was whoever could write the file's choice (R17-S03). Generous
    ///     against <see cref="MaxDebts" /> entries of ordinary paths.
    /// </remarks>
    private const long MaxQueueBytes = 8L * 1024 * 1024;

    /// <summary>Longest path a debt may name, so one entry cannot carry a megabyte.</summary>
    private const int MaxDebtPathLength = 4096;

    /// <summary>Longest error text kept with a debt.</summary>
    private const int MaxDebtErrorLength = 2048;

    /// <summary>How long a debt is retried before it is abandoned and reported.</summary>
    public static readonly TimeSpan DefaultRetention = TimeSpan.FromHours(24);

    /// <summary>Shortest wait between attempts.</summary>
    private static readonly TimeSpan BaseBackoff = TimeSpan.FromSeconds(30);

    /// <summary>Longest wait between attempts, however many have failed.</summary>
    private static readonly TimeSpan MaximumBackoff = TimeSpan.FromHours(1);

    /// <summary>
    ///     How paths are compared against the allowlist.
    ///     <para>
    ///         Case-insensitive only on Windows. macOS was included on the assumption that its
    ///         volumes are case-insensitive, but APFS and HFS+ can both be formatted
    ///         case-sensitive — and on such a volume two paths differing only in case are two
    ///         different files, so an ignore-case comparison would place one under a root that
    ///         does not contain it, and the ancestor walk would stop at a directory it never
    ///         reached (§22.7). The OS name does not determine the volume's semantics, so it is
    ///         not used to guess them: everywhere but Windows compares exactly.
    ///     </para>
    /// </summary>
    private static readonly StringComparison PathComparison = ComparisonFor(OperatingSystem.IsWindows());

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    /// <summary>
    ///     The lock objects handed out per canonical queue path, held weakly.
    ///     <para>
    ///         Strongly, this only ever grew: one lock and one key per distinct path, for the life
    ///         of the process, with nothing to remove them (R13-F03). A queue holds its own gate in
    ///         <see cref="_gate" /> for as long as it exists, so an entry whose target has been
    ///         collected is one no queue is using — and a queue that does not exist cannot be raced
    ///         with. Sharing is therefore unaffected; only the corpse is dropped.
    ///     </para>
    /// </summary>
    private static readonly Dictionary<string, WeakReference<object>> Gates =
        new(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);

    /// <summary>Guards <see cref="Gates" /> itself.</summary>
    private static readonly object GatesLock = new();

    private readonly IReadOnlyList<string> _allowedBasePaths;

    /// <summary>
    ///     The lock for this queue's file, shared by every instance pointing at it.
    ///     <para>
    ///         A per-instance lock is not a lock on the record: two hosts configured with the same
    ///         temp directory each held their own, so both could read, sweep and rewrite the same
    ///         JSON at once and one write could drop what the other had just recorded (§23.5).
    ///         Keyed on the canonical path, so queues in different directories never wait on each
    ///         other.
    ///     </para>
    /// </summary>
    private readonly object _gate;

    /// <summary>
    ///     The file whose handle keeps two processes out of the queue at the same time.
    ///     <para>
    ///         <see cref="_gate" /> orders the threads of one process and nothing more. Two servers
    ///         sharing a temp directory share this queue — which is the point of it, since a debt
    ///         has to outlive the process that recorded it — and both read the same list, each
    ///         added its own debt, and whichever wrote second erased the other's (R18-ARCH02).
    ///     </para>
    /// </summary>
    private readonly string _lockFile;

    private readonly string _queueFile;

    /// <summary>
    ///     Where this host's records live and the key they are signed with.
    ///     <para>
    ///         Held rather than read from a process-wide property: the queue and the journals of two
    ///         hosts in one process are two sets of records with two keys, and one global answer
    ///         signed each host's with the other's (R18-ARCH01).
    ///     </para>
    /// </summary>
    private readonly RecoveryContext _recovery;

    private readonly TimeSpan _retention;

    /// <summary>Creates a queue backed by a file.</summary>
    /// <param name="queueFile">Where the queue is kept between runs.</param>
    /// <param name="allowedBasePaths">Roots a deletion may happen under.</param>
    /// <param name="recovery">
    ///     Where this host's records live and the key they are signed with. Held rather than read
    ///     from a process-wide property, so two hosts in one process do not sign each other's
    ///     queues (R18-ARCH01).
    /// </param>
    /// <param name="retention">How long a debt is retried before being abandoned.</param>
    public CleanupDebtQueue(string queueFile, IReadOnlyList<string> allowedBasePaths,
        RecoveryContext recovery, TimeSpan? retention = null)
    {
        _queueFile = queueFile;
        _allowedBasePaths = allowedBasePaths;
        _recovery = recovery;
        _retention = retention ?? DefaultRetention;
        _gate = GateFor(queueFile);
        _lockFile = queueFile + ".lock";
    }

    /// <summary>How many paths the registry is tracking. A fixture's only view of its size.</summary>
    internal static int TrackedGates
    {
        get
        {
            lock (GatesLock)
            {
                return Gates.Count;
            }
        }
    }

    /// <summary>How long to wait for another process to finish with the queue.</summary>
    /// <remarks>A seam, so a fixture need not wait ten seconds to see a gate refused.</remarks>
    internal TimeSpan LockTimeout { get; set; } = CrossProcessFileGate.DefaultTimeout;

    /// <summary>The clock, so a fixture can drive retention without waiting for it.</summary>
    internal Func<DateTimeOffset> Now { get; set; } = () => DateTimeOffset.UtcNow;

    /// <summary>
    ///     Deletes one file. A seam, because the failure this queue exists for — a delete that
    ///     fails and then starts succeeding — cannot be produced from a fixture otherwise.
    /// </summary>
    internal Func<string, Func<FileStream, bool>, bool> DeleteIf { get; set; } =
        HandleBoundDelete.DeleteIf;

    /// <summary>
    ///     Writes the queue file. A seam, because a queue that cannot be persisted is the case
    ///     that most needs reporting and cannot be produced from a fixture otherwise.
    /// </summary>
    internal Action<Stream, string> WriteText { get; set; } = (stream, content) =>
    {
        using var writer = new StreamWriter(stream, leaveOpen: true);
        writer.Write(content);
    };

    /// <summary>
    ///     Told about anything that stops the queue from doing its job: a record it could not
    ///     read, one it could not write, or a debt it had no room for.
    ///     <para>
    ///         These used to be swallowed. A corrupt queue file read as an empty one and was then
    ///         overwritten, so a record of debts became no record at all with nothing said; a full
    ///         queue dropped the newest debt in silence. A queue whose whole purpose is that
    ///         nothing is quietly forgotten cannot itself forget quietly (§21.2.2).
    ///     </para>
    /// </summary>
    public Action<string>? OnQueueProblem { get; set; }

    /// <summary>
    ///     The comparison for a platform, as a function of that one fact.
    ///     <para>
    ///         Extracted so both answers can be checked on any machine. The rule was previously
    ///         only observable on the platform running the tests, so "we have no case-sensitive
    ///         volume to try this on" left the whole rule unverified rather than just its
    ///         deployment (§23.13.1). What a case-sensitive filesystem does with two names is not
    ///         in question; what this code does with them now is.
    ///     </para>
    /// </summary>
    /// <param name="isWindows">Whether this is the one platform that folds case in paths.</param>
    /// <returns>The comparison to use.</returns>
    internal static StringComparison ComparisonFor(bool isWindows)
    {
        return isWindows ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    }

    /// <summary>
    ///     Whether a path lies under a root, judged with a given platform's rule.
    ///     <para>
    ///         The containment decision itself, separated from the ambient platform so the
    ///         case-sensitive answer is testable where no case-sensitive volume exists.
    ///     </para>
    /// </summary>
    /// <param name="path">The canonical path to place.</param>
    /// <param name="root">The canonical root it should be under.</param>
    /// <param name="comparison">The platform's path comparison.</param>
    /// <returns><c>true</c> when the path is inside the root under that rule.</returns>
    internal static bool IsUnderRoot(string path, string root, StringComparison comparison)
    {
        var bounded = root.EndsWith(Path.DirectorySeparatorChar)
            ? root
            : root + Path.DirectorySeparatorChar;

        return path.StartsWith(bounded, comparison);
    }

    /// <summary>The lock shared by every queue on this file.</summary>
    /// <param name="queueFile">The queue's path.</param>
    /// <returns>The lock object for that path.</returns>
    private static object GateFor(string queueFile)
    {
        var key = Canonical(queueFile) ?? queueFile;

        lock (GatesLock)
        {
            if (Gates.TryGetValue(key, out var existing) && existing.TryGetTarget(out var shared))
                return shared;

            if (Gates.Count >= GateSweepThreshold) SweepCollectedGates();

            var gate = new object();
            Gates[key] = new WeakReference<object>(gate);
            return gate;
        }
    }

    /// <summary>Drops the entries whose lock nothing holds any more.</summary>
    /// <remarks>Called under <see cref="GatesLock" />.</remarks>
    private static void SweepCollectedGates()
    {
        var dead = Gates
            .Where(entry => !entry.Value.TryGetTarget(out _))
            .Select(entry => entry.Key)
            .ToList();

        foreach (var key in dead) Gates.Remove(key);
    }

    /// <summary>What a file looked like when a debt about it was recorded.</summary>
    /// <param name="path">The file.</param>
    /// <returns>Its creation time and length, or nulls when it cannot be read.</returns>
    /// <remarks>
    ///     The identity a replayed record has to match. A signature says this server recorded a
    ///     debt about that path once; it says nothing about what is at the path now, so a kept
    ///     record replayed after the name was reused deleted whoever had taken it (R18-SEC04).
    ///     Creation time and length are what .NET exposes portably — a file ID would be better and
    ///     needs a platform call — and together they make an accidental match unlikely and a
    ///     deliberate one require control of both.
    /// </remarks>
    private static (DateTimeOffset? CreatedUtc, long? Length, string? Digest) IdentityOf(
        string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

            // Everything from the one handle, so the length that decides whether to digest is the
            // length of the file being digested.
            var length = stream.Length;
            var created = File.GetCreationTimeUtc(path);

            if (length > MaxDigestBytes) return (created, length, null);

            var digest = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
            return (created, length, digest);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return (null, null, null);
        }
    }

    /// <summary>The text a debt's signature covers: the entry with its signature removed.</summary>
    /// <param name="debt">The debt to describe.</param>
    /// <returns>The canonical content.</returns>
    private static string SignedContent(Debt debt)
    {
        return JsonSerializer.Serialize(debt with { Signature = null }, Json);
    }

    /// <summary>Reads the queue, tolerating an absent or unreadable file.</summary>
    /// <returns>The debts currently recorded.</returns>
    /// <remarks>
    ///     Not gated across processes, unlike the two operations that write. Every write replaces
    ///     the file by an atomic rename, so a reader sees the list before or the list after and
    ///     never half of either; waiting for the gate would only make it a slightly later list.
    /// </remarks>
    public IReadOnlyList<Debt> Pending()
    {
        lock (_gate)
        {
            return Read();
        }
    }

    /// <summary>
    ///     Records a file a publish could not delete. Recording the same path again refreshes its
    ///     error without restarting its retention: the debt is as old as the first failure.
    /// </summary>
    /// <param name="path">The path the publish could not remove.</param>
    /// <param name="error">Why it could not.</param>
    /// <remarks>
    ///     Does nothing but report through <see cref="OnQueueProblem" /> when another process
    ///     holds the queue for longer than <see cref="LockTimeout" />. Writing anyway would erase
    ///     whatever that process is part-way through recording (R18-ARCH02).
    /// </remarks>
    public void Record(string path, string error)
    {
        var canonical = Canonical(path);
        if (canonical == null) return;

        lock (_gate)
        {
            using var scope = CrossProcessFileGate.TryAcquire(_lockFile, LockTimeout);
            if (scope == null)
            {
                OnQueueProblem?.Invoke(
                    $"The cleanup debt queue at '{_queueFile}' is held by another process, so "
                    + $"'{canonical}' was not recorded. Writing without the gate would erase "
                    + "whatever that process is in the middle of recording.");
                return;
            }

            var (debts, readable) = ReadState();
            if (!readable)
                // Writing now would replace an unreadable record with a one-entry one, losing
                // whatever it held. The problem has already been reported.
                return;

            var capability = _recovery.Capability;
            if (capability == null)
            {
                OnQueueProblem?.Invoke(
                    "No recovery signing key could be established, so a cleanup debt cannot be "
                    + $"recorded in a form the sweep will act on. '{canonical}' was not recorded.");
                return;
            }

            var existing = debts.FindIndex(d => string.Equals(d.Path, canonical, PathComparison));

            // Bounded before it is signed. The read side refuses a debt whose error is longer
            // than `MaxDebtErrorLength`, so an unbounded error signed in here was a debt the next
            // start silently dropped (R20-REC11).
            error = BoundedError(error);

            var now = Now();
            if (existing >= 0)
            {
                debts[existing] = Sign(debts[existing] with { LastError = error }, capability);
            }
            else if (debts.Count < MaxDebts)
            {
                var (createdUtc, length, digest) = IdentityOf(canonical);
                debts.Add(Sign(
                    new Debt(canonical, 0, now, now, error, createdUtc, length, digest),
                    capability));
            }
            else
            {
                OnQueueProblem?.Invoke(
                    $"The cleanup debt queue is full at {MaxDebts:N0} entries, so '{canonical}' "
                    + "was not recorded and will not be retried. An operator needs to look at the "
                    + "volume this queue is tracking.");
                return;
            }

            Write(debts);
        }
    }

    /// <summary>
    ///     Tries every debt that is due, and reports what happened to each.
    /// </summary>
    /// <returns>
    ///     What the sweep did, or four empty lists when the queue could not be read or another
    ///     process held it for longer than <see cref="LockTimeout" /> (R18-ARCH02). The debts are
    ///     still there either way; the next sweep picks them up.
    /// </returns>
    public SweepResult Sweep()
    {
        List<string> deleted = [], retrying = [], abandoned = [], refused = [];

        lock (_gate)
        {
            using var scope = CrossProcessFileGate.TryAcquire(_lockFile, LockTimeout);
            if (scope == null)
            {
                // Another process has the queue. It will sweep what it holds; doing it here as
                // well would mean two processes deleting from one list and writing back two
                // different remainders.
                OnQueueProblem?.Invoke(
                    $"The cleanup debt queue at '{_queueFile}' is held by another process, so this "
                    + "sweep did nothing. The next one will pick the debts up.");
                return new SweepResult([], [], [], []);
            }

            var now = Now();
            var remaining = new List<Debt>();

            var (recorded, readable) = ReadState();
            if (!readable) return new SweepResult([], [], [], []);

            foreach (var debt in recorded)
            {
                if (debt.NextAttemptUtc > now)
                {
                    remaining.Add(debt);
                    continue;
                }

                var eligibility = ClassifyDeletion(debt.Path);
                if (eligibility != DeletionEligibility.Allowed)
                {
                    refused.Add(debt.Path);
                    if (eligibility == DeletionEligibility.OutsideThisHost)
                        // This host cannot decide the debt's expiry: another host sharing the
                        // queue may have an allowlist that covers it (R20-REC06, R21-REC06).
                        remaining.Add(debt);

                    continue;
                }

                if (!File.Exists(debt.Path))
                {
                    deleted.Add(debt.Path);
                    continue;
                }

                // Expiry before the deletion, not after a failed one. A record kept past its
                // retention was still acted on as long as the delete succeeded, so age was no
                // barrier to replaying an old one (R18-SEC04).
                if (now - debt.FirstSeenUtc >= _retention)
                {
                    abandoned.Add(debt.Path);
                    continue;
                }

                // Judged through the handle that deletes. The identity was checked by pathname
                // and then `HandleBoundDelete` reopened the pathname, so the digest that said
                // "this is the file" and the handle that removed it need not have been about the
                // same object (R20-REC02). A debt with no recorded identity is refused outright.
                if (!HasARecordedIdentity(debt))
                {
                    refused.Add(debt.Path);
                    continue;
                }

                try
                {
                    if (!DeleteIf(debt.Path, stream => IsTheRecordedFile(debt, stream)))
                    {
                        // The name is now somebody else's file. That is not this host's debt any
                        // more, and keeping it would delete whoever takes the name next.
                        refused.Add(debt.Path);
                        continue;
                    }

                    deleted.Add(debt.Path);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    var attempts = debt.Attempts + 1;
                    if (now - debt.FirstSeenUtc >= _retention)
                    {
                        abandoned.Add(debt.Path);
                        continue;
                    }

                    // Re-signed, because the signature covers the whole entry: a debt whose
                    // attempt count and next-attempt time have moved on is a different record, and
                    // keeping the old signature would make the next sweep drop it as unverified.
                    remaining.Add(Resign(debt with
                    {
                        Attempts = attempts,
                        NextAttemptUtc = now + Backoff(attempts),
                        LastError = ex.Message
                    }));
                    retrying.Add(debt.Path);
                }
            }

            Write(remaining);
        }

        return new SweepResult(deleted, retrying, abandoned, refused);
    }

    /// <summary>How long to wait after a given number of failures.</summary>
    /// <param name="attempts">Failures so far.</param>
    /// <returns>The wait before the next attempt.</returns>
    private static TimeSpan Backoff(int attempts)
    {
        var doublings = Math.Min(attempts, 10);
        var wait = TimeSpan.FromTicks(BaseBackoff.Ticks * (1L << doublings));
        return wait > MaximumBackoff ? MaximumBackoff : wait;
    }

    /// <summary>An error string cut to what a debt may carry.</summary>
    /// <param name="error">The error as reported.</param>
    /// <returns>The error, no longer than <see cref="MaxDebtErrorLength" />.</returns>
    private static string BoundedError(string error)
    {
        return error.Length <= MaxDebtErrorLength ? error : error[..MaxDebtErrorLength];
    }

    /// <summary>Whether a debt carries the identity the sweep needs to judge its target.</summary>
    /// <param name="debt">The debt.</param>
    /// <returns><c>true</c> when it records both a length and a digest.</returns>
    /// <remarks>
    ///     A record with no identity is one written before this existed. Those are refused rather
    ///     than trusted: a debt whose target cannot be confirmed is exactly the shape a replay
    ///     takes (R18-SEC04). A target too large to digest was recorded without one and is refused
    ///     for the same reason — deleting on evidence already shown to be forgeable is the thing
    ///     being stopped.
    /// </remarks>
    private static bool HasARecordedIdentity(Debt debt)
    {
        return debt is { TargetLength: not null, TargetDigest: not null };
    }

    /// <summary>Whether the open file is the one the debt was recorded about.</summary>
    /// <param name="debt">The debt.</param>
    /// <param name="stream">The file, open on the handle that will delete it.</param>
    /// <returns><c>true</c> when its length and digest are the recorded ones.</returns>
    /// <remarks>
    ///     Length and digest only. The recorded creation time is a pathname question and is not
    ///     asked here: the point of judging on the handle is that nothing between the judgement
    ///     and the delete can be about a different object, and a creation time read by name would
    ///     be exactly that (R20-REC02).
    ///     <para>
    ///         The content decides it, not the metadata. A creation time and a length are both the
    ///         local actor's to choose — NTFS hands a recreated file its predecessor's creation
    ///         time without being asked, and padding to a length is arithmetic — so a replacement
    ///         made to match them was deleted as though it were the original (R19-REC05). The
    ///         length is still compared first, because a mismatch there saves reading the file.
    ///     </para>
    /// </remarks>
    private static bool IsTheRecordedFile(Debt debt, FileStream stream)
    {
        return debt.TargetLength == stream.Length
               && debt.TargetDigest != null
               && string.Equals(debt.TargetDigest, HandleBoundDelete.DigestOf(stream),
                   StringComparison.Ordinal);
    }

    /// <summary>
    ///     Classifies whether this host may delete the path now, judged from the filesystem as it
    ///     is.
    /// </summary>
    /// <param name="path">The recorded path.</param>
    /// <returns>
    ///     <see cref="DeletionEligibility.Allowed" /> when the target is inside this host's
    ///     allowlist and contains no links; <see cref="DeletionEligibility.OutsideThisHost" />
    ///     when another host may cover it; otherwise <see cref="DeletionEligibility.UnsafePath" />.
    /// </returns>
    private DeletionEligibility ClassifyDeletion(string path)
    {
        var canonical = Canonical(path);
        if (canonical == null) return DeletionEligibility.UnsafePath;
        if (!string.Equals(canonical, path, PathComparison)) return DeletionEligibility.UnsafePath;

        string? root = null;
        if (_allowedBasePaths.Count > 0)
        {
            root = _allowedBasePaths
                .Select(Canonical)
                .FirstOrDefault(candidate => candidate != null && IsUnder(canonical, candidate));
            if (root == null) return DeletionEligibility.OutsideThisHost;
        }

        return NothingOnThePathIsALink(canonical, root)
            ? DeletionEligibility.Allowed
            : DeletionEligibility.UnsafePath;
    }

    /// <summary>Whether a canonical path lies under a canonical root.</summary>
    /// <param name="path">The path to place.</param>
    /// <param name="root">The root it should be under.</param>
    /// <returns><c>true</c> when the path is inside the root.</returns>
    private static bool IsUnder(string path, string root)
    {
        return IsUnderRoot(path, root, PathComparison);
    }

    /// <summary>
    ///     Whether the file and every directory between it and the allowlisted root is a real
    ///     entry rather than a reparse point.
    ///     <para>
    ///         Checking only the leaf was not containment: an ancestor directory replaced by a
    ///         junction after the debt was recorded redirects the whole path out of the root while
    ///         the leaf itself is a perfectly ordinary file (§21.2.2). A junction needs no
    ///         elevation on Windows, so this is a step an unprivileged local process can take.
    ///     </para>
    /// </summary>
    /// <param name="canonical">The canonical path about to be deleted.</param>
    /// <param name="root">The allowlisted root to walk up to, or null to walk to the volume.</param>
    /// <returns><c>true</c> when no component of the path is a link.</returns>
    /// <remarks>
    ///     Shared with <see cref="PublishJournal" />, which asked the same question of the leaf
    ///     alone and so accepted a destination reached through a junction. Two copies of a rule
    ///     that differ is the shape this codebase keeps paying for, so there is one.
    /// </remarks>
    internal static bool NothingOnThePathIsALink(string canonical, string? root)
    {
        try
        {
            var file = new FileInfo(canonical);
            if (file is { Exists: true, LinkTarget: not null }) return false;

            for (var directory = file.Directory; directory != null; directory = directory.Parent)
            {
                if (directory is { Exists: true, LinkTarget: not null }) return false;

                if (root != null
                    && string.Equals(directory.FullName.TrimEnd(Path.DirectorySeparatorChar),
                        root.TrimEnd(Path.DirectorySeparatorChar), PathComparison))
                    return true;
            }

            // The walk reached the volume without meeting the root, which means the path is not
            // under it however the strings compared.
            return root == null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or SecurityException)
        {
            return false;
        }
    }

    /// <summary>The fully resolved form of a path, or null when it cannot be resolved.</summary>
    /// <param name="path">The path to resolve.</param>
    /// <returns>The canonical path, or null.</returns>
    private static string? Canonical(string path)
    {
        try
        {
            return string.IsNullOrWhiteSpace(path) ? null : Path.GetFullPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException
                                       or PathTooLongException or IOException)
        {
            return null;
        }
    }

    /// <summary>
    ///     Reads the queue file, distinguishing "there is nothing recorded" from "the record
    ///     cannot be read".
    /// </summary>
    /// <returns>The recorded debts, and whether the file was readable.</returns>
    private (List<Debt> Debts, bool Readable) ReadState()
    {
        try
        {
            if (!File.Exists(_queueFile)) return ([], true);

            // Size and content on one handle. This read the whole file whatever its size (R17-S03),
            // and then bounded it through a `FileInfo` and read it through a second open, so the
            // file the cap was measured on need not be the file that was read (R18-SEC03).
            if (!SecureFile.TryReadBounded(_queueFile, MaxQueueBytes, out var text, out var length))
            {
                OnQueueProblem?.Invoke(
                    $"The cleanup debt queue at '{_queueFile}' is at least {length:N0} bytes, larger "
                    + $"than a queue may be ({MaxQueueBytes:N0}), so it was not read.");
                return ([], false);
            }

            var stored = JsonSerializer.Deserialize<List<Debt>>(text, Json) ?? [];

            if (stored.Count > MaxDebts)
            {
                OnQueueProblem?.Invoke(
                    $"The cleanup debt queue at '{_queueFile}' holds {stored.Count:N0} entries, "
                    + $"more than the {MaxDebts:N0} this records, so it was not read.");
                return ([], false);
            }

            return (stored.Where(IsWellFormed).ToList(), true);
        }
        catch (Exception ex) when (ex is IOException or JsonException
                                       or UnauthorizedAccessException)
        {
            // Not empty — unreadable. Returning an empty list here made the next write replace a
            // record of real debts with nothing at all (§21.2.2).
            OnQueueProblem?.Invoke(
                $"The cleanup debt queue at '{_queueFile}' could not be read, so the debts it "
                + $"holds cannot be retried or reported: {ex.Message}");
            return ([], false);
        }
    }

    /// <summary>Whether a stored debt is one this queue could have written.</summary>
    /// <param name="debt">The entry as it was read.</param>
    /// <returns><c>true</c> when it is well formed and signed by this installation.</returns>
    /// <remarks>
    ///     Shape before anything else, and both before the filesystem is asked about the path. A
    ///     UNC path handed to <c>FileInfo</c> is an SMB request on Windows, and the refusal that
    ///     followed could not take the request back (R17-S04). And a well-formed path is still only
    ///     a path: what makes it an instruction is that this installation signed it (R17-S01).
    /// </remarks>
    private bool IsWellFormed(Debt debt)
    {
        if (string.IsNullOrWhiteSpace(debt.Path)) return false;
        if (debt.Path.Length > MaxDebtPathLength) return false;
        if (debt.LastError.Length > MaxDebtErrorLength) return false;
        if (PublishJournal.LooksRemoteOrSpecial(debt.Path)) return false;
        if (!Path.IsPathFullyQualified(debt.Path)) return false;

        var capability = _recovery.Capability;
        return capability != null && capability.Verify(SignedContent(debt), debt.Signature);
    }

    /// <summary>The recorded debts, or an empty list when the record cannot be read.</summary>
    /// <returns>The recorded debts.</returns>
    private List<Debt> Read()
    {
        return ReadState().Debts;
    }

    /// <summary>Signs a debt with this installation's key, when there is one.</summary>
    /// <param name="debt">The debt to sign.</param>
    /// <returns>The signed debt, or the debt unchanged when no key could be established.</returns>
    /// <remarks>
    ///     An unsigned debt is dropped on the next read, which is the right outcome: without a key
    ///     nothing on disk can be told from anything else on disk, so nothing should be acted on.
    /// </remarks>
    private Debt Resign(Debt debt)
    {
        return _recovery.Capability is { } capability
            // Bounded here, on every path that re-signs. `Record` bounded the error it was
            // handed, and the retry path then signed `ex.Message` unbounded; the next read
            // refused the whole debt as malformed and the cleanup was forgotten (R21-REC01).
            ? Sign(debt with { LastError = BoundedError(debt.LastError) }, capability)
            : debt;
    }

    /// <summary>Signs a debt so the sweep will act on it.</summary>
    /// <param name="debt">The debt to sign.</param>
    /// <param name="capability">This installation's signing key.</param>
    /// <returns>The debt with its signature.</returns>
    private static Debt Sign(Debt debt, RecoveryCapability capability)
    {
        return debt with { Signature = capability.Sign(SignedContent(debt with { Signature = null })) };
    }

    /// <summary>
    ///     Writes the queue, through a staging file so a crash mid-write cannot leave a truncated
    ///     one behind.
    /// </summary>
    /// <param name="debts">The debts to record.</param>
    private void Write(List<Debt> debts)
    {
        try
        {
            var directory = Path.GetDirectoryName(_queueFile);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

            // A staging name carrying a nonce, not `<queue>.writing`: that name is predictable,
            // and a path-based write to a name someone has left a link at follows the link
            // (R18-SEC02).
            SecureFile.ReplaceAtomically(_queueFile, JsonSerializer.Serialize(debts, Json),
                (stream, content) => WriteText(stream, content));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A queue that cannot be written is worth no more than the publish it followed, and
            // that publish already succeeded. Failing the caller's completed operation because of
            // it would be worse — but it is said out loud rather than swallowed, because a debt
            // that was never persisted is one nobody will ever act on (§21.2.2).
            OnQueueProblem?.Invoke(
                $"The cleanup debt queue at '{_queueFile}' could not be written, so "
                + $"{debts.Count:N0} pending cleanup(s) will not survive a restart: {ex.Message}");
        }
    }

    /// <summary>How the current host should handle a recorded deletion target.</summary>
    private enum DeletionEligibility
    {
        /// <summary>The target is inside this host's allowlist and has no link components.</summary>
        Allowed,

        /// <summary>Another host may cover the target, so the debt must remain queued.</summary>
        OutsideThisHost,

        /// <summary>The path changed or became unsafe, so the debt must not be replayed later.</summary>
        UnsafePath
    }

    /// <summary>One file a publish replaced but could not remove.</summary>
    /// <param name="Path">The canonical path of the file still on disk.</param>
    /// <param name="Attempts">How many deletions have been tried and failed.</param>
    /// <param name="FirstSeenUtc">When the debt was first recorded, for retention.</param>
    /// <param name="NextAttemptUtc">The earliest time the next attempt may be made.</param>
    /// <param name="LastError">Why the last attempt failed.</param>
    /// <param name="TargetCreatedUtc">
    ///     When the file this debt is about was created, as it was when the debt was recorded.
    /// </param>
    /// <param name="TargetLength">Its length when the debt was recorded.</param>
    /// <param name="TargetDigest">
    ///     A SHA-256 of its content when the debt was recorded, or null when it was too large to
    ///     digest within <see cref="MaxDigestBytes" />.
    /// </param>
    /// <param name="Signature">
    ///     The HMAC over every other field, proving this installation recorded it. A debt that does
    ///     not verify is a path someone wrote into the queue file, not a cleanup this server asked
    ///     for (R17-S01).
    /// </param>
    public sealed record Debt(
        string Path,
        int Attempts,
        DateTimeOffset FirstSeenUtc,
        DateTimeOffset NextAttemptUtc,
        string LastError,
        DateTimeOffset? TargetCreatedUtc = null,
        long? TargetLength = null,
        string? TargetDigest = null,
        string? Signature = null);

    /// <summary>What one sweep did.</summary>
    /// <param name="Deleted">Paths removed this sweep.</param>
    /// <param name="Retrying">Paths that failed again and remain queued.</param>
    /// <param name="Abandoned">Paths that outlived the retention and need an operator.</param>
    /// <param name="Refused">
    ///     Paths the queue declined to act on — outside the allowlist, or now a reparse point.
    ///     These leave the queue: re-attempting a path that has become something else is how a
    ///     cleanup record turns into a deletion primitive.
    /// </param>
    public sealed record SweepResult(
        IReadOnlyList<string> Deleted,
        IReadOnlyList<string> Retrying,
        IReadOnlyList<string> Abandoned,
        IReadOnlyList<string> Refused);
}
