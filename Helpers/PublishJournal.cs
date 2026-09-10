using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Records a multi-file publish on disk while it is in flight, so a process that dies partway
///     leaves something the next start can put right.
///     <para>
///         <see cref="BoundedFileBatch" /> already rolls back within the process: a move that fails
///         restores every destination it had replaced. A <em>crash</em> is different — nothing runs
///         the rollback, and what is left is some destinations replaced, some backups sitting as
///         <c>.replaced-*</c> siblings, and no record of which was which. That gap has been on the
///         residual list since the tenth round.
///     </para>
///     <para>
///         The journal is written before the first move and deleted after the last one, so its mere
///         existence at startup means a publish did not finish. Recovery restores each backup over
///         its destination — the state the caller's request would have left had it been refused,
///         which is the outcome the in-process rollback also produces.
///     </para>
///     <para>
///         What it covers is a process that dies, not a machine that loses power. Neither the
///         journal nor the staging files are flushed to the device and no directory is fsynced, so
///         after a power loss a record can be missing steps that did reach the disk, or describe
///         steps that did not (R18-CONTRACT02). Making that claim true means <c>Flush(true)</c> on
///         every journal write and an fsync of the directory after each rename, at a cost this has
///         not been measured against; until then the weaker guarantee is the one stated.
///     </para>
///     <para>
///         Deliberately conservative about what it will act on. A journal names destinations and
///         backups that were canonicalised and allowlisted when it was written, and recovery
///         re-checks both before touching anything: a journal is a record of intent, exactly like
///         the cleanup queue, and never an authority to move files.
///     </para>
///     <para>
///         <b>A journal is only an instruction if this server wrote it.</b> Every record carries an
///         HMAC over its own content, keyed by a secret kept beside it
///         (<see cref="RecoveryCapability" />), and recovery ignores anything that does not verify.
///         Before that, authority came from the paths a record named — canonical, under a trusted
///         root, not a link — and containment is not provenance. A record naming a file inside an
///         allowed root was indistinguishable from one this server wrote (R17-S02).
///     </para>
///     <para>
///         Two earlier arguments were wrong and are worth keeping as the reason this is signed
///         rather than reasoned about. The first was that restoring a backup gives a forger
///         nothing, because planting
///         <c>&lt;destination&gt;.replaced-&lt;32 hex&gt;</c> means being able to write beside the
///         destination. The second was that a content digest makes deletion safe, because knowing
///         a file's exact bytes means being able to read it. Both stopped holding the moment
///         journals moved into one central recovery directory: a forger then needs to write only
///         <em>there</em>, not beside the victim (R17-S02). A signature does not depend on
///         reasoning about what else the attacker can reach.
///     </para>
///     <para>
///         <b>A finished publish is never undone.</b> The record's state is <c>Publishing</c> while
///         destinations are being replaced and <c>Committed</c> once they are all in place. A
///         committed record is only ever tidied up — backups removed, journal deleted — and its
///         destinations are not touched. Without that, a successful publish whose journal could not
///         be deleted was read at the next start as an interrupted one, and for a created step the
///         digest matches by construction, so the file that had just been delivered was removed
///         (R17-F01).
///     </para>
/// </summary>
public static class PublishJournal
{
    /// <summary>The suffix that marks a file as an in-flight publish record.</summary>
    public const string Extension = ".publishing.json";

    /// <summary>Most steps a journal may carry, so a crafted file cannot ask for unbounded work.</summary>
    private const int MaxSteps = 10_000;

    /// <summary>
    ///     Most journals one directory scan will process. A per-file size cap bounds each read and
    ///     said nothing about how many files there are; a directory can be filled with small,
    ///     well-formed journals to make start-up take as long as someone likes.
    /// </summary>
    private const int MaxJournalsPerScan = 512;

    /// <summary>
    ///     Largest journal this will read, in bytes. Read before the file is opened for content:
    ///     an unbounded read of an untrusted file in a scanned directory is work an attacker
    ///     chooses the size of (R13-SEC01).
    /// </summary>
    private const long MaxJournalBytes = 4L * 1024 * 1024;

    /// <summary>The name of the ledger of transactions recovery has already executed.</summary>
    public const string LedgerName = "recovered.ledger.json";

    /// <summary>The file whose handle claims a recovery root for one process at a time.</summary>
    public const string ClaimName = "recovery.claim";

    /// <summary>Maximum UTF-8 bytes accepted from one pending-journal marker.</summary>
    private const long MaxPendingMarkerBytes = 1024;

    /// <summary>The most transactions the ledger remembers.</summary>
    /// <remarks>
    ///     A ledger that grows without bound is a file an attacker can make this server read at
    ///     every start. Oldest entries fall off first: a transaction old enough to have been
    ///     pushed out has had every start since to be noticed, and its destinations have long
    ///     since been overwritten by ordinary use.
    /// </remarks>
    public const int MaxLedgerEntries = 4096;

    /// <summary>A transaction that is still replacing destinations.</summary>
    public const string Publishing = "publishing";

    /// <summary>A transaction whose destinations are all in place.</summary>
    public const string Committed = "committed";

    /// <summary>The most outstanding journals a new publish generation admits per root.</summary>
    internal const int JournalNameCeiling = 8192;

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    /// <summary>Test seam after index registration and before the journal appears.</summary>
    internal static Action? AfterIndexRegistrationForTest { get; set; }

    /// <summary>The SHA-256 of a file, lower-case hex, or null when it cannot be read.</summary>
    /// <param name="path">The file to hash.</param>
    /// <returns>The digest, or null.</returns>
    public static string? DigestOf(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return Convert.ToHexString(SHA256.HashData(stream))
                .ToLowerInvariant();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    /// <summary>The transaction id of a record already read, when it verifies.</summary>
    /// <param name="record">The record.</param>
    /// <param name="capability">This installation's signing key.</param>
    /// <returns>The id, or null when the record does not verify or has none.</returns>
    private static string? TransactionIdOf(Record? record, RecoveryCapability capability)
    {
        if (record?.TransactionId == null) return null;

        return capability.Verify(SignedContent(record), record.Signature)
            ? record.TransactionId
            : null;
    }

    /// <summary>Reads a journal once, within what the budget has left, and charges what it read.</summary>
    /// <param name="path">The journal.</param>
    /// <param name="budget">The start-up's budget; what is read is charged to it before this returns.</param>
    /// <returns>
    ///     The record (null when the file deserialises to nothing), the bytes read, a refusal when
    ///     it could not be read or was not affordable, and whether it was the budget that refused
    ///     it — in which case nothing after it is affordable either.
    /// </returns>
    /// <remarks>
    ///     Every journal used to be read twice — once for its id, once to recover — and charged
    ///     to the budget once, on a `FileInfo` length taken before either read (R21-REC03). One
    ///     read, and the number charged is the number of bytes that actually came in.
    ///     <para>
    ///         Read up to <c>min(MaxJournalBytes, remaining budget)</c>, decided on the length of
    ///         the handle being read (writers are excluded by the share mode, so that length is
    ///         the file's) before a byte comes in; and charged in <c>finally</c>, so a file that
    ///         fails to parse costs what it cost. Reading the whole file first and charging only
    ///         if it parsed let malformed journals do 4 MiB of work each for nothing on the
    ///         64 MiB aggregate bound (R22-REC01).
    ///     </para>
    /// </remarks>
    private static (Record? Record, string? Refusal, bool OverBudget) ReadJournal(
        string path, RecoveryBudget budget)
    {
        var cap = Math.Min(MaxJournalBytes, Math.Max(0, budget.Bytes));
        var read = 0L;
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var length = stream.Length;
            if (length > MaxJournalBytes)
                return (null, $"{path}: {length:N0} bytes, larger than a journal may be", false);
            if (length > cap)
                return (null,
                    $"{path}: {length:N0} bytes, more than this start-up's remaining recovery budget "
                    + $"of {budget.Bytes:N0}; the next start will begin here", true);

            var buffer = new byte[length];
            int step;
            while (read < length && (step = stream.Read(buffer, (int)read, (int)(length - read))) > 0)
                read += step;

            var text = Encoding.UTF8.GetString(buffer, 0, (int)read);
            return (JsonSerializer.Deserialize<Record>(text, Json), null, false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or JsonException)
        {
            return (null, $"{path}: could not be read ({ex.Message})", false);
        }
        finally
        {
            budget.Charge(read);
        }
    }

    /// <summary>Removes a journal whose transaction was already carried out.</summary>
    /// <param name="path">The journal file.</param>
    /// <param name="refused">Where to report a removal that failed.</param>
    /// <remarks>
    ///     Left in place it would be read again at every start for as long as it exists, and the
    ///     ledger would have to keep its id for just as long. Removing it is safe precisely
    ///     because the ledger says its work is done.
    /// </remarks>
    private static void TidyReplayed(string path, List<string> refused)
    {
        try
        {
            File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            refused.Add($"{path}: an already-recovered record could not be removed ({ex.Message})");
        }
    }

    /// <summary>Transactions this installation has already recovered.</summary>
    /// <param name="directory">The recovery directory the ledger lives in.</param>
    /// <param name="capability">The key the ledger is signed with.</param>
    /// <returns>The transaction ids, or an empty set when there is no ledger this may trust.</returns>
    /// <remarks>
    ///     Signed like every other record here, and for the same reason: an unsigned ledger is
    ///     one anyone could empty, which would turn the replay defence off rather than on.
    /// </remarks>
    internal static Ledger? ReadLedger(string directory, RecoveryCapability capability)
    {
        var path = Path.Combine(directory, LedgerName);

        try
        {
            // No ledger at all is the ordinary first start. A ledger that is present but cannot be
            // verified is something else entirely: read as empty, it switched the replay defence
            // off, which is the one thing a tampered ledger would be for (R20-REC09). Null here is
            // the caller's signal to suspend rather than proceed.
            if (!File.Exists(path)) return new Ledger([]);
            if (!SecureFile.TryReadBounded(path, MaxJournalBytes, out var text, out _)) return null;

            var ledger = JsonSerializer.Deserialize<Ledger>(text, Json);
            if (ledger?.Recovered == null || ledger.Signature == null) return null;

            return capability.Verify(LedgerContent(ledger.Recovered, ledger.Generation),
                ledger.Signature)
                ? ledger
                : null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or JsonException)
        {
            return null;
        }
    }

    /// <summary>Records that a transaction has been recovered, so it cannot be replayed.</summary>
    /// <param name="directory">The recovery directory.</param>
    /// <param name="capability">The key to sign with.</param>
    /// <param name="recovered">Every transaction id now known to have been executed.</param>
    /// <param name="generation">The write count, one more than the ledger this replaces.</param>
    /// <param name="stillPresent">
    ///     Ids whose journal is still in the directory. Kept ahead of everything else when the
    ///     ledger is over its cap, because a tombstone for a journal that is gone protects nothing,
    ///     and one for a journal still there is the only thing stopping its replay (R20-REC09).
    /// </param>
    /// <param name="absenceProven">
    ///     Whether <paramref name="stillPresent" /> is complete — every journal was read. When it
    ///     is not, nothing is evicted (R21-REC02).
    /// </param>
    internal static void WriteLedger(string directory, RecoveryCapability capability,
        List<string> recovered, long generation, IReadOnlySet<string> stillPresent,
        bool absenceProven = true)
    {
        // Eviction only when every journal was seen. `stillPresent` is what this start-up read,
        // and a journal it never reached — over the count, over the bytes, past the deadline —
        // is not absent, so its tombstone must not be the one to go (R21-REC02). A ledger over
        // its cap after a truncated scan stays over it until a complete scan can evict honestly.
        if (recovered.Count > MaxLedgerEntries && absenceProven)
        {
            // Evict what can no longer be replayed first: an id whose journal is still in the
            // directory is exactly the one a tombstone exists for, and dropping it because it was
            // oldest reopened the replay (R20-REC09). Only then the oldest of the rest.
            var evictable = recovered.Where(id => !stillPresent.Contains(id)).ToList();
            var keep = recovered.Where(stillPresent.Contains).ToList();
            var room = Math.Max(0, MaxLedgerEntries - keep.Count);
            recovered = evictable.Skip(Math.Max(0, evictable.Count - room)).Concat(keep).ToList();
        }

        try
        {
            SecureFile.ReplaceAtomically(Path.Combine(directory, LedgerName),
                JsonSerializer.Serialize(
                    new Ledger(recovered, capability.Sign(LedgerContent(recovered, generation)),
                        generation), Json));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Nothing to do but leave it: the next start re-reads whatever is there, and a ledger
            // that could not be written means a transaction may be recovered twice rather than
            // never. Reported by the caller through the refusal list.
        }
    }

    /// <summary>The text a ledger's signature covers: its generation and every id.</summary>
    /// <param name="recovered">The transaction ids.</param>
    /// <param name="generation">The write count.</param>
    /// <returns>The canonical content.</returns>
    private static string LedgerContent(IReadOnlyList<string> recovered, long generation)
    {
        return generation + "\n" + string.Join("\n", recovered);
    }

    /// <summary>The text a record's signature covers: every field except the signature.</summary>
    /// <param name="record">The record to describe.</param>
    /// <returns>The canonical content.</returns>
    /// <remarks>
    ///     Serialised from the record with the signature removed, so the signed text is exactly
    ///     what is stored and there is no second format to drift from the first.
    /// </remarks>
    private static string SignedContent(Record record)
    {
        return JsonSerializer.Serialize(record with { Signature = null }, Json);
    }

    /// <summary>Signs a record so recovery will act on it.</summary>
    /// <param name="record">The record to sign.</param>
    /// <param name="capability">This installation's signing key.</param>
    /// <returns>The record with its signature.</returns>
    public static Record Signed(Record record, RecoveryCapability capability)
    {
        // Given here rather than by the caller, so no path through this class can produce a record
        // recovery cannot tell apart from a replay of another (R19-REC03). The signature covers it,
        // because an id an attacker may edit identifies nothing.
        var identified = record with
        {
            TransactionId = record.TransactionId ?? Guid.NewGuid().ToString("N")
        };

        return identified with { Signature = capability.Sign(SignedContent(identified)) };
    }

    /// <summary>A record describing what this process is about to do.</summary>
    /// <param name="what">What the batch is publishing.</param>
    /// <param name="steps">The destinations it is working on.</param>
    /// <param name="state">
    ///     Whether the transaction is still replacing destinations or has committed.
    /// </param>
    /// <returns>The record, stamped with this process's identity.</returns>
    public static Record Describe(string what, IReadOnlyList<Entry> steps, string state)
    {
        using var self = Process.GetCurrentProcess();
        return new Record(what, steps, Environment.MachineName, self.Id,
            new DateTimeOffset(self.StartTime.ToUniversalTime(), TimeSpan.Zero), state);
    }

    /// <summary>Whether the process that wrote a journal is still running.</summary>
    /// <param name="record">The journal's record.</param>
    /// <returns><c>true</c> when that publish may still be in flight.</returns>
    /// <remarks>
    ///     A journal exists while its publish runs, so one written by a live process is not a
    ///     crash — it is a transaction in progress, and rolling it back would corrupt the thing
    ///     this exists to protect. Unknown answers "no": a journal from another host, or from a
    ///     process this one may not query, is treated as recoverable, because the alternative
    ///     lets an unanswerable question block recovery for good.
    /// </remarks>
    private static bool ItsWriterIsStillRunning(Record record)
    {
        if (record.ProcessId is not { } id || record.StartedUtc is not { } started) return false;
        if (!string.Equals(record.Machine, Environment.MachineName, StringComparison.OrdinalIgnoreCase))
            return false;

        try
        {
            using var process = Process.GetProcessById(id);
            var actual = new DateTimeOffset(process.StartTime.ToUniversalTime(), TimeSpan.Zero);
            return Math.Abs((actual - started).TotalSeconds) < 1;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException
                                       or NotSupportedException or SystemException)
        {
            return false;
        }
    }

    /// <summary>Writes or replaces a journal.</summary>
    /// <param name="path">Where the journal lives.</param>
    /// <param name="record">What the publish is doing.</param>
    /// <remarks>
    ///     This is the direct single-record API paired with
    ///     <see cref="Recover(string,IReadOnlyList{string},RecoveryCapability,Action{string,string})" />.
    ///     Production batch
    ///     publishing uses <see cref="WriteManaged" /> so startup discovery is registered before
    ///     the first destination mutation.
    /// </remarks>
    public static void Write(string path, Record record)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

        // Staged then moved, so a crash during the write cannot leave a half-written journal that
        // recovery would then read as a truncated list of steps — through a staging name carrying a
        // nonce, because a predictable one can have a link planted at it and a path-based write
        // follows the link (R18-SEC02).
        SecureFile.ReplaceAtomically(path, JsonSerializer.Serialize(record, Json));
    }

    /// <summary>Registers and writes a journal that startup recovery must discover.</summary>
    /// <param name="path">Where the managed journal lives.</param>
    /// <param name="record">What the publish is doing.</param>
    /// <exception cref="IOException">
    ///     Thrown when the marker index cannot be trusted, admission is full, or the journal write
    ///     fails. The caller invokes this before its first destination mutation.
    /// </exception>
    internal static void WriteManaged(string path, Record record)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        WriteIndexedJournal(path, JsonSerializer.Serialize(record, Json));
    }

    /// <summary>Removes a journal once its publish has finished.</summary>
    /// <param name="path">The journal to remove, or null when the batch wrote none.</param>
    public static void Delete(string? path)
    {
        TryDelete(path);
    }

    /// <summary>Removes a journal, saying whether it is actually gone.</summary>
    /// <param name="path">The journal, or null when the batch wrote none.</param>
    /// <returns><c>true</c> when nothing is left at the path.</returns>
    /// <remarks>
    ///     The answer matters to the ledger and to nobody else. A publish that finished does not
    ///     care whether its journal could be removed — recovery finds nothing to do with it later —
    ///     but a transaction may only be recorded as executed once its journal is provably gone,
    ///     because that record is what stops the next start from touching it (R20-REC01).
    /// </remarks>
    private static bool TryDelete(string? path)
    {
        if (path == null) return true;

        try
        {
            if (File.Exists(path)) File.Delete(path);
            return !File.Exists(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    ///     Puts back what an unfinished publish had already replaced, and removes the journal.
    /// </summary>
    /// <param name="path">The journal to recover from.</param>
    /// <param name="trustedRoots">
    ///     The roots recovery may act under, from the caller's configuration. Never read from the
    ///     journal: that made the document being validated its own authority (R13-SEC01).
    /// </param>
    /// <param name="capability">
    ///     This installation's signing key. A record that does not verify is not an instruction,
    ///     so it is refused before the filesystem is asked about anything it names (R17-S02).
    /// </param>
    /// <param name="restore">
    ///     Moves a backup back over its destination. A seam: the failure that matters here cannot
    ///     be produced from a fixture otherwise.
    /// </param>
    /// <returns>What was restored, what was refused, and what was left behind.</returns>
    public static RecoveryResult Recover(string path, IReadOnlyList<string> trustedRoots,
        RecoveryCapability capability, Action<string, string>? restore = null)
    {
        // Bounded on the handle the bound is measured on (R18-SEC03). Unreadable is not empty:
        // the journal is left in place and the refusal says so, rather than deleting the only
        // record that a publish was interrupted.
        var (record, refusal, _) = ReadJournal(path, new RecoveryBudget { Bytes = MaxJournalBytes });
        if (refusal != null) return new RecoveryResult([], [refusal], []);

        return Recover(path, record, trustedRoots, capability, restore);
    }

    /// <summary>Recovers from a record already read, so a start-up reads each journal once.</summary>
    /// <param name="path">The journal, for deletion and messages.</param>
    /// <param name="record">Its content, or null when the file deserialised to nothing.</param>
    /// <param name="trustedRoots">The roots a restore may act under.</param>
    /// <param name="capability">This installation's signing key.</param>
    /// <param name="restore">A seam for the move, or null for <see cref="File.Move(string, string, bool)" />.</param>
    /// <returns>What was restored and what was refused.</returns>
    internal static RecoveryResult Recover(string path, Record? record,
        IReadOnlyList<string> trustedRoots, RecoveryCapability capability,
        Action<string, string>? restore)
    {
        List<string> restored = [], refused = [], orphaned = [];

        if (record == null || record.Steps.Count == 0
                           || record.Steps.Count > MaxSteps)
        {
            Delete(path);
            return new RecoveryResult([], [], []);
        }

        // Before anything else, and before any path in it reaches the filesystem: a record this
        // server did not write is not an instruction, whatever the paths in it look like.
        if (!capability.Verify(SignedContent(record), record.Signature))
            // Left in place. Deleting it would let anyone who can write here remove a real
            // interrupted publish's record by planting an unsigned file over its name.
            return new RecoveryResult([],
                [$"{path}: not signed by this installation, so it is not acted on"], []);

        // A journal whose writer is still running describes a publish in flight, not a crash.
        // Rolling it back mid-transaction would corrupt exactly what this protects — and two hosts
        // sharing a temp directory is a configuration the debt sink stack already anticipates.
        if (ItsWriterIsStillRunning(record))
            return new RecoveryResult([], [$"{path}: the publish that wrote it is still running"], []);

        // Every destination was in place when this was written, so the transaction succeeded and
        // only its leftovers remain. Touching a destination here would undo a delivered result —
        // and for a created step the digest matches by construction, so "roll it back" would have
        // meant "delete exactly the file that was just published" (R17-F01).
        if (string.Equals(record.State, Committed, StringComparison.Ordinal))
            return TidyCommitted(path, record, trustedRoots);

        // Structure first, filesystem second. `File.Exists` used to run before any of this, so a
        // crafted UNC destination produced a network lookup before anything refused it — the
        // lookup *was* the effect (R13-SEC01).
        var malformed = record.Steps
            .Where(step => !IsStructurallyValid(step))
            .Select(step => step.Destination)
            .ToList();

        if (malformed.Count > 0)
        {
            Delete(path);
            return new RecoveryResult([],
                malformed.Select(name => $"{name}: not a well-formed journal entry").ToList(), []);
        }

        // The restore is a rename, and a rename is by name: .NET offers no handle-bound move, so
        // the backup and destination are resolved here and again by `File.Move`. The checks above
        // run immediately before it, which is the narrowest that window can be made without a
        // platform rename-by-handle; it is the weaker guarantee, stated (R20-REC03).
        var move = restore ?? ((from, to) => File.Move(from, to, true));

        // Reverse order, like the in-process rollback: the last destination replaced is the first
        // put back, so a destination touched twice ends as it began.
        foreach (var step in record.Steps.Reverse())
        {
            // No backup means this publish *created* the destination, and rolling that back
            // means removing it. Removed only when the bytes on disk are the ones the journal says
            // the publish wrote: naming a file is a request to delete anything, naming its
            // contents can only ask for the file the publish actually produced. Anything else is
            // reported and left. See the class summary.
            if (step.Backup == null)
            {
                if (!File.Exists(step.Destination)) continue;

                var reasonForCreated = WhyNotRecoverable(step, trustedRoots);
                if (reasonForCreated != null)
                {
                    refused.Add($"{step.Destination}: {reasonForCreated}");
                    continue;
                }

                if (step.Created == null)
                {
                    orphaned.Add(step.Destination);
                    continue;
                }

                // Judged and removed on one handle. `DigestOf(path)` followed by `File.Delete(path)`
                // resolved the name twice, so the bytes that matched the journal and the file that
                // was deleted need not have been the same object (R20-REC03).
                try
                {
                    if (HandleBoundDelete.DeleteIf(step.Destination,
                            stream => string.Equals(step.Created, HandleBoundDelete.DigestOf(stream),
                                StringComparison.Ordinal)))
                        restored.Add(step.Destination);
                    else
                        orphaned.Add(step.Destination);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    refused.Add($"{step.Destination}: could not be removed ({ex.Message})");
                }

                continue;
            }

            // A backup that is no longer there means the publish reached its commit point and
            // tidied it: the journal is stale, not actionable. Counting that as a refusal kept the
            // journal forever and reported a problem where there was none.
            if (!File.Exists(step.Backup)) continue;

            var reason = WhyNotRecoverable(step, trustedRoots);
            if (reason != null)
            {
                refused.Add($"{step.Destination}: {reason}");
                continue;
            }

            try
            {
                move(step.Backup, step.Destination);
                restored.Add(step.Destination);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                refused.Add($"{step.Destination}: could not be restored ({ex.Message})");
            }
        }

        // Kept when anything was refused, so the next start can try again rather than losing the
        // record of an incomplete publish. Orphans do not keep it: nothing further will be done
        // about them, and holding the journal would report them at every start for ever.
        var consumed = refused.Count == 0 && TryDelete(path);

        return new RecoveryResult(restored, refused, orphaned, consumed);
    }

    /// <summary>
    ///     Clears up after a transaction that finished: its backups, and its journal.
    /// </summary>
    /// <param name="path">The journal.</param>
    /// <param name="record">Its verified content.</param>
    /// <param name="trustedRoots">The roots a backup may be removed from.</param>
    /// <returns>What was tidied and what was refused.</returns>
    /// <remarks>
    ///     No destination is read or written. The only files touched are the backups this
    ///     transaction itself created, whose names are the destination's own plus
    ///     <c>.replaced-</c> and a nonce, and which the entry-shape rules already require.
    /// </remarks>
    private static RecoveryResult TidyCommitted(string path, Record record,
        IReadOnlyList<string> trustedRoots)
    {
        List<string> removed = [], refused = [];

        foreach (var step in record.Steps)
        {
            if (step.Backup == null || !IsStructurallyValid(step)) continue;
            if (!File.Exists(step.Backup)) continue;

            var reason = WhyNotRecoverable(step, trustedRoots);
            if (reason != null)
            {
                refused.Add($"{step.Backup}: {reason}");
                continue;
            }

            try
            {
                File.Delete(step.Backup);
                removed.Add(step.Backup);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                refused.Add($"{step.Backup}: could not be removed ({ex.Message})");
            }
        }

        var consumed = refused.Count == 0 && TryDelete(path);

        return new RecoveryResult([], refused, removed, consumed);
    }

    /// <summary>
    ///     Whether an entry is well formed, decided without touching the filesystem.
    /// </summary>
    /// <param name="step">The entry to check.</param>
    /// <returns><c>true</c> when it is safe to ask the filesystem about.</returns>
    /// <remarks>
    ///     Everything here is a property of the strings themselves. A UNC or device path is
    ///     refused before any lookup, and a backup must be the sibling this code itself would have
    ///     created — same directory, the destination's name, and the `.replaced-` marker. An
    ///     attacker choosing both halves of a move cannot name an arbitrary pair.
    /// </remarks>
    private static bool IsStructurallyValid(Entry? step)
    {
        if (step == null || string.IsNullOrWhiteSpace(step.Destination)) return false;
        if (LooksRemoteOrSpecial(step.Destination)) return false;
        if (!Path.IsPathFullyQualified(step.Destination)) return false;

        if (step.Backup == null) return true;

        if (string.IsNullOrWhiteSpace(step.Backup)) return false;
        if (LooksRemoteOrSpecial(step.Backup)) return false;
        if (!Path.IsPathFullyQualified(step.Backup)) return false;

        // The shape this code writes: "<destination>.replaced-<32 hex>".
        var expectedPrefix = step.Destination + ".replaced-";
        if (!step.Backup.StartsWith(expectedPrefix, StringComparison.Ordinal)) return false;

        var nonce = step.Backup[expectedPrefix.Length..];
        return nonce.Length == 32 && nonce.All(Uri.IsHexDigit);
    }

    /// <summary>Whether a path is a form this server never writes to.</summary>
    /// <param name="path">The path as the journal recorded it.</param>
    /// <returns><c>true</c> when it is a UNC, device or stream path.</returns>
    internal static bool LooksRemoteOrSpecial(string path)
    {
        if (path.StartsWith(@"\\", StringComparison.Ordinal)) return true;
        if (path.StartsWith("//", StringComparison.Ordinal)) return true;
        if (path.Contains(@"\\?\", StringComparison.Ordinal)) return true;
        if (path.Contains(@"\\.\", StringComparison.Ordinal)) return true;

        // An alternate data stream is named after the colon that follows the drive letter.
        var afterRoot = path.Length > 2 ? path[2..] : string.Empty;
        return afterRoot.Contains(':', StringComparison.Ordinal);
    }

    /// <summary>
    ///     Whether a journal entry may be acted on, judged from the filesystem as it is now.
    /// </summary>
    /// <param name="step">The entry to check.</param>
    /// <param name="trustedRoots">
    ///     The roots recovery may act under, from the caller's configuration — never from the
    ///     journal (R13-SEC01).
    /// </param>
    /// <returns>The reason it may not be acted on, or null when it may.</returns>
    private static string? WhyNotRecoverable(Entry step, IReadOnlyList<string> trustedRoots)
    {
        // A created step has no backup; asking about a null path would throw where a refusal is
        // wanted.
        var paths = step.Backup == null
            ? new[] { step.Destination }
            : new[] { step.Destination, step.Backup };

        foreach (var path in paths)
        {
            string canonical;
            try
            {
                canonical = Path.GetFullPath(path);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException
                                           or PathTooLongException or IOException)
            {
                return "its path could not be resolved";
            }

            if (!string.Equals(canonical, path, StringComparison.Ordinal))
                return "its path no longer resolves to what the journal recorded";

            // No roots configured means the server may write anywhere, which is the default. That
            // used to be read as "no authority" and refused everything — correct when the paths
            // were all recovery had to judge by, and wrong once the record is signed: it left the
            // default configuration with a crash recovery that recovered nothing (R17-F02). The
            // signature is the authority; the roots are a second bound where one is configured.
            if (trustedRoots.Count > 0 && !trustedRoots.Any(root => IsUnder(canonical, root)))
                return "it is outside every root this server may write to";
        }

        // The same rule the cleanup queue applies, and now literally the same code. Two things
        // were wrong with the version here. It read `Backup ?? Destination`, so in the restore
        // branch the destination — the file being written *over* — was never examined at all
        // (independent review). And it asked only about the leaf: an ancestor directory replaced
        // by a junction redirects the whole path out of the trusted root while the file at the end
        // of it is perfectly ordinary, which is the containment failure §21.2.2 fixed in the queue
        // and never brought here.
        return paths.Any(candidate => !CleanupDebtQueue.NothingOnThePathIsALink(candidate,
            trustedRoots.FirstOrDefault(root => IsUnder(candidate, root))))
            ? "it, or a directory on the way to it, has become a link"
            : null;
    }

    /// <summary>Whether a canonical path lies under a root.</summary>
    /// <param name="path">The path to place.</param>
    /// <param name="root">The root it should be under.</param>
    /// <returns><c>true</c> when the path is inside the root.</returns>
    private static bool IsUnder(string path, string root)
    {
        var canonical = Path.GetFullPath(root);
        var bounded = canonical.EndsWith(Path.DirectorySeparatorChar)
            ? canonical
            : canonical + Path.DirectorySeparatorChar;

        return CleanupDebtQueue.IsUnderRoot(path, bounded.TrimEnd(Path.DirectorySeparatorChar),
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
    }

    /// <summary>
    ///     Recovers every unfinished publish left in a directory.
    /// </summary>
    /// <param name="directory">Where journals are kept.</param>
    /// <param name="trustedRoots">
    ///     The roots recovery may act under, from configuration rather than from any
    ///     journal it finds.
    /// </param>
    /// <param name="capability">
    ///     This installation's signing key. A record it did not produce is not an instruction.
    /// </param>
    /// <returns>What was restored and what was refused, across all of them.</returns>
    public static RecoveryResult RecoverAll(string directory, IReadOnlyList<string> trustedRoots,
        RecoveryCapability capability)
    {
        // The single-directory form: this directory is its own state directory, and the claim is
        // taken here. The host's start-up takes one claim itself and calls the form below per root.
        var budget = new RecoveryBudget();
        using var claim = CrossProcessFileGate.TryAcquire(Path.Combine(directory, ClaimName),
            budget.Remaining);
        if (claim == null)
            return new RecoveryResult([],
                [$"{directory}: another process is recovering it, so this start did not"], []);

        return RecoverAll(directory, trustedRoots, capability, budget, directory);
    }

    /// <summary>Recovers one root, spending from a budget shared across the whole start-up.</summary>
    /// <param name="directory">The directory to scan.</param>
    /// <param name="trustedRoots">The roots recovery may act under.</param>
    /// <param name="capability">This installation's signing key.</param>
    /// <param name="budget">What this start-up has left to spend.</param>
    /// <param name="stateDirectory">
    ///     Where the ledger lives: the host's private recovery directory, whichever root is being
    ///     scanned. Writing it into the scanned root put recovery state into an operator's output
    ///     directory whenever that root was a legacy allowlisted one (R20-REC10). The caller holds
    ///     the claim on this directory for the whole start-up.
    /// </param>
    /// <returns>What was restored and what was refused.</returns>
    public static RecoveryResult RecoverAll(string directory, IReadOnlyList<string> trustedRoots,
        RecoveryCapability capability, RecoveryBudget budget, string stateDirectory)
    {
        if (budget.Exhausted)
            return new RecoveryResult([],
            [
                $"{directory}: this start-up's recovery budget was already spent, so it was not "
                + "scanned; the next start will begin here"
            ], []);

        List<string> restored = [], refused = [], orphaned = [];

        // The claim is the caller's (R19-REC04, R20-REC10): one per start-up, on the state
        // directory, so a legacy root is never written to and a second root does not re-take a
        // gate this process already holds.
        var ledger = ReadLedger(stateDirectory, capability);
        if (ledger == null)
            return new RecoveryResult([],
            [
                $"{stateDirectory}: the recovery ledger is present but cannot be verified, so "
                + "no interrupted publish will be recovered until an operator removes or restores "
                + "it; acting on journals without a trustworthy ledger is how a replay gets in"
            ], []);

        var alreadyRecovered = ledger.Recovered.ToList();

        List<string> journals;
        bool inventoryComplete;
        var cursorPath = Path.Combine(stateDirectory, CursorNameFor(directory));
        try
        {
            // The persistent marker index is deterministically bucketed, so recovery resumes
            // after the actual last attempt without rebuilding or truncating an opaque filesystem
            // prefix. Legacy flat journals are inventoried once; managed writers register before
            // the journal appears and new work is refused at the admission ceiling.
            var cursor = ReadCursor(cursorPath);
            var requested = budget.Journals == int.MaxValue
                ? int.MaxValue
                : Math.Max(0, budget.Journals) + 1;
            var inventory = ReadPendingNames(directory, stateDirectory, cursor, requested,
                budget.Remaining);
            inventoryComplete = inventory.Complete;
            journals = inventory.Names
                .Select(name => Path.Combine(directory, name))
                .ToList();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new RecoveryResult([], [$"{directory}: could not be read ({ex.Message})"], []);
        }

        // Whether every journal in this directory was looked at. Only then is "not in `present`"
        // the same as "gone", which is what eviction needs (R21-REC02). Decided before the list
        // is cut down to what is affordable, or a cut-down list would always look complete. The
        // persistent index contains every future-generation name. A legacy bootstrap may contain
        // more than the admission ceiling, but bounded bucket pages and the cursor still traverse
        // the whole inventory across starts.
        var affordable = budget.Take(journals.Count);
        var scanComplete = inventoryComplete && journals.Count <= affordable;
        if (journals.Count > affordable)
        {
            journals = journals.Take(affordable).ToList();
            refused.Add(
                $"{directory}: more publish records than this start-up's remaining budget; the "
                + "rest were left for a later start rather than spending this one on them");
        }

        var executed = new List<string>(alreadyRecovered);
        var present = new HashSet<string>(StringComparer.Ordinal);

        string? lastAttempted = null;
        foreach (var journal in journals)
            try
            {
                // A genuine record that has already been carried out is not an instruction to
                // carry it out again: kept and put back, it restored backups over destinations
                // that had moved on since (R19-REC03).
                // Charged before it is read: the bytes budget is what bounds the work an
                // unauthenticated file can ask for (R20-RES01). Over budget stops the scan here.
                // Spent before this one? Then this one is not read at all. Bytes and time only:
                // this root's share of the count was taken above, for all of these at once.
                if (budget.BytesOrTimeSpent)
                {
                    refused.Add($"{journal}: this start-up's recovery budget was spent before it; "
                                + "the next start will begin here");
                    scanComplete = false;
                    break;
                }

                lastAttempted = Path.GetFileName(journal);

                // One read, bounded by what is left and charged at what it cost, inside
                // `ReadJournal`: a journal that exactly uses up the budget is affordable, one byte
                // more is not read at all (R21-REC04, R22-REC01). Whether anything is left is the
                // next journal's question.
                var (record, readRefusal, overBudget) = ReadJournal(journal, budget);
                if (overBudget)
                {
                    refused.Add(readRefusal!);
                    scanComplete = false;
                    break;
                }

                if (readRefusal != null)
                {
                    // Present, but not read: its identity is unknown, so it cannot be in
                    // `present`, and treating "not in present" as "gone" would let eviction drop
                    // the tombstone of a journal still sitting here (R22-REC04).
                    refused.Add(readRefusal);
                    scanComplete = false;
                    continue;
                }

                var identity = TransactionIdOf(record, capability);
                if (identity != null) present.Add(identity);
                if (identity != null && alreadyRecovered.Contains(identity))
                {
                    refused.Add($"{journal}: this transaction has already been recovered");
                    TidyReplayed(journal, refused);
                    continue;
                }

                var result = Recover(journal, record, trustedRoots, capability, null);
                restored.AddRange(result.Restored);
                refused.AddRange(result.Refused);
                orphaned.AddRange(result.Orphaned);

                // Only a journal that is gone is executed. One kept for a retry is still an
                // instruction, and writing its id down here made the next start read the retry
                // as a replay and delete it (R20-REC01).
                if (identity != null && result.Consumed) executed.Add(identity);
            }
            catch (Exception ex)
            {
                // One unreadable or hostile journal must not stop the rest of this root from
                // being recovered — a single crafted file would otherwise deny recovery to every
                // legitimate one beside it (R13-SEC01). It is still here, though, and whatever it
                // is, its absence is not proven (R22-REC04).
                refused.Add($"{journal}: recovery failed ({ex.GetType().Name}: {ex.Message})");
                scanComplete = false;
            }

        if (executed.Count != alreadyRecovered.Count)
            WriteLedger(stateDirectory, capability, executed, ledger.Generation + 1, present,
                scanComplete);

        // Where to continue from: the last name actually attempted, not the last one selected
        // before a byte/deadline stop. A failed write is operator-visible; silently returning to
        // the first retained record on every start made the same liveness bug reappear below the
        // inventory ceiling.
        var cursorWritten = scanComplete
            ? WriteCursor(cursorPath, null)
            : lastAttempted == null || WriteCursor(cursorPath, lastAttempted);
        if (!cursorWritten)
            refused.Add($"{cursorPath}: recovery progress could not be persisted; later starts "
                        + "may repeat this batch until the state directory is writable");

        return new RecoveryResult(restored, refused, orphaned);
    }

    /// <summary>Registers and writes a discoverable journal under one index gate.</summary>
    /// <param name="path">The journal path.</param>
    /// <param name="json">The serialised signed record.</param>
    /// <exception cref="IOException">
    ///     Thrown when the index cannot be locked, verified or updated, or the root is at capacity.
    /// </exception>
    private static void WriteIndexedJournal(string path, string json)
    {
        var fullPath = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(fullPath)
                        ?? throw new IOException("A publish journal has no containing directory.");
        var name = Path.GetFileName(fullPath);
        if (!IsJournalName(name))
            throw new IOException($"'{name}' is not a publish-journal file name.");

        using var gate = CrossProcessFileGate.TryAcquire(IndexGatePath(directory, directory));
        if (gate == null)
            throw new IOException($"The pending-journal index for '{directory}' is busy.");

        EnsureIndexBootstrapped(directory, directory);
        if (!MarkerExists(directory, directory, name))
        {
            var pending = ReadPendingNamesUnderGate(directory, directory, null,
                JournalNameCeiling);
            if (pending.Names.Count >= JournalNameCeiling)
                throw new IOException(
                    $"The recovery root already holds at least {pending.Names.Count:N0} pending "
                    + $"publish journals; the {JournalNameCeiling:N0}-journal safety limit was "
                    + "reached before any destination was changed.");
            WriteMarker(directory, directory, name);
        }

        // Recovery takes the same gate before pruning missing reservations. Keeping it until the
        // journal appears prevents a concurrent scan from deleting this reservation in the only
        // window where it is intentionally missing.
        AfterIndexRegistrationForTest?.Invoke();
        SecureFile.ReplaceAtomically(fullPath, json);
    }

    /// <summary>Reads every currently pending name from the guarded inventory.</summary>
    /// <param name="directory">The recovery root whose journals are indexed.</param>
    /// <param name="stateDirectory">The private directory that holds this root's index.</param>
    /// <param name="cursor">The last journal attempted on the previous start, or null.</param>
    /// <param name="limit">The most present names to return in this page.</param>
    /// <param name="timeout">How long remains in the startup budget.</param>
    /// <returns>A bounded page and whether it completed one full index rotation.</returns>
    /// <exception cref="IOException">Thrown when the index cannot be locked or trusted.</exception>
    private static PendingInventory ReadPendingNames(string directory, string stateDirectory,
        string? cursor, int limit, TimeSpan timeout)
    {
        using var gate = CrossProcessFileGate.TryAcquire(IndexGatePath(directory, stateDirectory),
            timeout);
        if (gate == null)
            throw new IOException($"The pending-journal index for '{directory}' is busy.");

        EnsureIndexBootstrapped(directory, stateDirectory);
        return ReadPendingNamesUnderGate(directory, stateDirectory, cursor, limit);
    }

    /// <summary>Creates the persistent marker index for pre-index flat journals once.</summary>
    /// <param name="directory">The recovery root whose journals are indexed.</param>
    /// <param name="stateDirectory">The private directory holding the index.</param>
    /// <exception cref="IOException">Thrown when the index cannot be created or verified.</exception>
    private static void EnsureIndexBootstrapped(string directory, string stateDirectory)
    {
        var ready = IndexReadyPath(directory, stateDirectory);
        if (File.Exists(ready))
        {
            if (!SecureFile.TryReadBounded(ready, 8, out var text, out _)
                || !string.Equals(text, "1", StringComparison.Ordinal))
                throw new IOException("The pending-journal bootstrap marker is not valid.");
            if (!Directory.Exists(IndexRootPath(directory, stateDirectory)))
                throw new IOException(
                    "The pending-journal marker index is missing after bootstrap completed.");
            EnsurePlainDirectory(IndexRootPath(directory, stateDirectory));
            return;
        }

        EnsurePlainDirectory(IndexRootPath(directory, stateDirectory));
        if (Directory.Exists(directory))
            foreach (var journal in Directory.EnumerateFiles(directory, "*" + Extension,
                         SearchOption.TopDirectoryOnly))
            {
                var name = Path.GetFileName(journal);
                if (IsJournalName(name)) WriteMarker(directory, stateDirectory, name);
            }

        SecureFile.ReplaceAtomically(ready, "1");
    }

    /// <summary>Reads one bounded rotation from a bootstrapped marker index.</summary>
    /// <param name="directory">The recovery root whose journals are indexed.</param>
    /// <param name="stateDirectory">The private directory holding the index.</param>
    /// <param name="cursor">The last journal attempted on the previous start, or null.</param>
    /// <param name="limit">The most present names to return.</param>
    /// <returns>A bounded page and whether it completed one full index rotation.</returns>
    /// <exception cref="IOException">Thrown when marker state cannot be verified or pruned.</exception>
    private static PendingInventory ReadPendingNamesUnderGate(string directory,
        string stateDirectory, string? cursor, int limit)
    {
        if (limit <= 0) return new PendingInventory([], false);

        var root = IndexRootPath(directory, stateDirectory);
        EnsurePlainDirectory(root);
        var buckets = Directory.EnumerateDirectories(root, "*", SearchOption.TopDirectoryOnly)
            .Select(path => (Path: path, Name: Path.GetFileName(path)))
            .ToList();
        var invalidBucket = buckets.FirstOrDefault(bucket => !IsBucketName(bucket.Name));
        if (invalidBucket.Path != null)
            throw new IOException(
                $"The pending-journal index contains an invalid bucket '{invalidBucket.Name}'.");

        var byName = buckets.ToDictionary(bucket => bucket.Name, bucket => bucket.Path,
            StringComparer.Ordinal);
        var orderedBuckets = byName.Keys.OrderBy(name => name, StringComparer.Ordinal).ToList();
        var segments = new List<(string Bucket, bool AfterCursor)>();
        if (cursor == null)
        {
            segments.AddRange(orderedBuckets.Select(bucket => (bucket, false)));
        }
        else
        {
            var cursorBucket = BucketName(cursor);
            if (byName.ContainsKey(cursorBucket)) segments.Add((cursorBucket, true));
            segments.AddRange(orderedBuckets.Where(bucket =>
                string.CompareOrdinal(bucket, cursorBucket) > 0).Select(bucket => (bucket, false)));
            segments.AddRange(orderedBuckets.Where(bucket =>
                string.CompareOrdinal(bucket, cursorBucket) < 0).Select(bucket => (bucket, false)));
            if (byName.ContainsKey(cursorBucket)) segments.Add((cursorBucket, false));
        }

        var selected = new List<string>();
        Dictionary<string, List<(string Name, string Marker)>> cache = new(StringComparer.Ordinal);
        foreach (var segment in segments)
        {
            if (!cache.TryGetValue(segment.Bucket, out var entries))
            {
                entries = ReadBucket(directory, byName[segment.Bucket]);
                cache[segment.Bucket] = entries;
            }

            IEnumerable<(string Name, string Marker)> part = entries;
            if (cursor != null && string.Equals(segment.Bucket, BucketName(cursor),
                    StringComparison.Ordinal))
                part = segment.AfterCursor
                    ? entries.Where(entry => string.CompareOrdinal(entry.Name, cursor) > 0)
                    : entries.Where(entry => string.CompareOrdinal(entry.Name, cursor) <= 0);

            foreach (var entry in part)
            {
                selected.Add(entry.Name);
                if (selected.Count == limit) return new PendingInventory(selected, false);
            }
        }

        return new PendingInventory(selected, true);
    }

    /// <summary>Reads and validates every marker in one small deterministic bucket.</summary>
    /// <param name="directory">The recovery root containing the journals.</param>
    /// <param name="bucketDirectory">The private marker bucket.</param>
    /// <returns>Present names and their marker paths in ordinal name order.</returns>
    /// <exception cref="IOException">Thrown when a marker is invalid or cannot be pruned.</exception>
    private static List<(string Name, string Marker)> ReadBucket(string directory,
        string bucketDirectory)
    {
        EnsurePlainDirectory(bucketDirectory);
        var present = new List<(string Name, string Marker)>();
        foreach (var marker in Directory.EnumerateFiles(bucketDirectory, "*.entry",
                     SearchOption.TopDirectoryOnly))
        {
            if (!SecureFile.TryReadBounded(marker, MaxPendingMarkerBytes, out var name, out var length))
                throw new IOException(
                    $"The pending-journal marker '{marker}' is at least {length:N0} bytes, above "
                    + $"its {MaxPendingMarkerBytes:N0}-byte bound.");
            if (!IsJournalName(name)
                || !string.Equals(Path.GetFileName(marker), MarkerFileName(name),
                    StringComparison.Ordinal))
                throw new IOException($"The pending-journal marker '{marker}' is not valid.");

            if (File.Exists(Path.Combine(directory, name)))
            {
                present.Add((name, marker));
                continue;
            }

            File.Delete(marker);
            if (File.Exists(marker))
                throw new IOException($"The stale pending-journal marker '{marker}' could not be removed.");
        }

        return present.OrderBy(entry => entry.Name, StringComparer.Ordinal).ToList();
    }

    /// <summary>Writes one idempotent marker for a managed or bootstrapped journal.</summary>
    /// <param name="directory">The recovery root containing the journal.</param>
    /// <param name="stateDirectory">The private directory holding the index.</param>
    /// <param name="name">The journal file name.</param>
    /// <exception cref="IOException">Thrown when an existing marker does not match.</exception>
    private static void WriteMarker(string directory, string stateDirectory, string name)
    {
        var bucket = Path.Combine(IndexRootPath(directory, stateDirectory), BucketName(name));
        EnsurePlainDirectory(bucket);
        var marker = Path.Combine(bucket, MarkerFileName(name));
        if (File.Exists(marker))
        {
            if (!SecureFile.TryReadBounded(marker, MaxPendingMarkerBytes, out var existing, out _)
                || !string.Equals(existing, name, StringComparison.Ordinal))
                throw new IOException($"The pending-journal marker '{marker}' does not match its name.");
            return;
        }

        SecureFile.ReplaceAtomically(marker, name);
    }

    /// <summary>Whether the exact journal already has a verified marker.</summary>
    /// <param name="directory">The recovery root containing the journal.</param>
    /// <param name="stateDirectory">The private directory holding the index.</param>
    /// <param name="name">The journal file name.</param>
    /// <returns><c>true</c> only for a present marker containing that exact name.</returns>
    /// <exception cref="IOException">Thrown when a present marker cannot be verified.</exception>
    private static bool MarkerExists(string directory, string stateDirectory, string name)
    {
        var marker = MarkerPath(directory, stateDirectory, name);
        if (!File.Exists(marker)) return false;
        if (!SecureFile.TryReadBounded(marker, MaxPendingMarkerBytes, out var existing, out _)
            || !string.Equals(existing, name, StringComparison.Ordinal))
            throw new IOException($"The pending-journal marker '{marker}' does not match its name.");
        return true;
    }

    /// <summary>Whether a pending-index entry is a file name in the journal namespace.</summary>
    /// <param name="name">The candidate file name.</param>
    /// <returns>Whether the value is one base name carrying the exact journal suffix.</returns>
    private static bool IsJournalName(string? name)
    {
        return !string.IsNullOrEmpty(name)
               && string.Equals(Path.GetFileName(name), name, StringComparison.Ordinal)
               && name.EndsWith(Extension, StringComparison.Ordinal);
    }

    /// <summary>The marker-index directory for one recovery root.</summary>
    /// <param name="directory">The recovery root.</param>
    /// <param name="stateDirectory">The private state directory.</param>
    /// <returns>The root of the deterministic marker buckets.</returns>
    private static string IndexRootPath(string directory, string stateDirectory)
    {
        return Path.Combine(stateDirectory, $"journals.{DirectoryDigest(directory)}.pending");
    }

    /// <summary>The durable marker saying legacy inventory completed.</summary>
    /// <param name="directory">The recovery root.</param>
    /// <param name="stateDirectory">The private state directory.</param>
    /// <returns>The bootstrap marker path.</returns>
    private static string IndexReadyPath(string directory, string stateDirectory)
    {
        return Path.Combine(stateDirectory, $"journals.{DirectoryDigest(directory)}.pending.ready");
    }

    /// <summary>The cross-process gate protecting one pending index.</summary>
    /// <param name="directory">The recovery root.</param>
    /// <param name="stateDirectory">The private state directory.</param>
    /// <returns>The cross-process gate path.</returns>
    private static string IndexGatePath(string directory, string stateDirectory)
    {
        return Path.Combine(stateDirectory, $"journals.{DirectoryDigest(directory)}.pending.claim");
    }

    /// <summary>The exact marker path for one journal name.</summary>
    /// <param name="directory">The recovery root.</param>
    /// <param name="stateDirectory">The private state directory.</param>
    /// <param name="name">The journal file name.</param>
    /// <returns>The marker path.</returns>
    private static string MarkerPath(string directory, string stateDirectory, string name)
    {
        return Path.Combine(IndexRootPath(directory, stateDirectory), BucketName(name),
            MarkerFileName(name));
    }

    /// <summary>An ordinal-order-preserving bucket for a journal name.</summary>
    /// <param name="name">The journal file name.</param>
    /// <returns>The bucket name.</returns>
    private static string BucketName(string name)
    {
        return string.Concat(name.Take(12)
            .Select(character => ((int)character).ToString("x4")));
    }

    /// <summary>The collision-resistant marker file name for a journal name.</summary>
    /// <param name="name">The journal file name.</param>
    /// <returns>The marker file name.</returns>
    private static string MarkerFileName(string name)
    {
        return NameDigest(name) + ".entry";
    }

    /// <summary>A stable SHA-256 digest of one ordinal journal name.</summary>
    /// <param name="name">The journal file name.</param>
    /// <returns>Lower-case hexadecimal SHA-256.</returns>
    private static string NameDigest(string name)
    {
        return Convert.ToHexString(
                SHA256.HashData(Encoding.UTF8.GetBytes(name)))
            .ToLowerInvariant();
    }

    /// <summary>Whether a directory name is one lower-case twelve-code-unit bucket.</summary>
    /// <param name="name">The candidate name.</param>
    /// <returns>Whether it is a bucket name.</returns>
    private static bool IsBucketName(string name)
    {
        return name.Length == 48
               && name.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f');
    }

    /// <summary>Creates or validates an ordinary, non-link index directory.</summary>
    /// <param name="path">The directory path.</param>
    /// <exception cref="IOException">Thrown when the path is not an ordinary directory.</exception>
    private static void EnsurePlainDirectory(string path)
    {
        if (!Directory.Exists(path)) Directory.CreateDirectory(path);
        var info = new DirectoryInfo(path);
        info.Refresh();
        if (!info.Exists || (info.Attributes & FileAttributes.ReparsePoint) != 0)
            throw new IOException($"The pending-journal index directory '{path}' is not an ordinary directory.");
    }

    /// <summary>The cursor file for one root, in the state directory.</summary>
    /// <param name="directory">The root being recovered.</param>
    /// <returns>A file name unique to the root.</returns>
    private static string CursorNameFor(string directory)
    {
        return "journals." + DirectoryDigest(directory) + ".cursor";
    }

    /// <summary>A stable, non-secret identifier for a recovery root.</summary>
    /// <param name="directory">The recovery root.</param>
    /// <returns>A lower-case hexadecimal prefix unique under platform path semantics.</returns>
    private static string DirectoryDigest(string directory)
    {
        var canonical = Path.GetFullPath(directory);
        if (OperatingSystem.IsWindows()) canonical = canonical.ToUpperInvariant();
        var digest = SHA256.HashData(
            Encoding.UTF8.GetBytes(canonical));
        return Convert.ToHexString(digest)[..16].ToLowerInvariant();
    }

    /// <summary>Reads the last journal name a previous start reached in a root.</summary>
    /// <param name="path">The cursor file.</param>
    /// <returns>The name, or null when there is no usable cursor.</returns>
    /// <exception cref="IOException">Thrown when a present cursor cannot be trusted.</exception>
    private static string? ReadCursor(string path)
    {
        try
        {
            if (!File.Exists(path)) return null;
            var name = File.ReadAllText(path).Trim();
            if (name.Length == 0) return null;
            if (IsJournalName(name)) return name;
            throw new IOException(
                $"The recovery cursor '{path}' is not a publish-journal file name. Restore or "
                + "remove it before recovery resumes for this root.");
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new IOException(
                $"The recovery cursor '{path}' could not be read. Recovery is suspended for "
                + "this root until the cursor is repaired.", ex);
        }
    }

    /// <summary>Writes down the last journal name this start reached, or clears it.</summary>
    /// <param name="path">The cursor file.</param>
    /// <param name="name">The name, or null once the root was covered.</param>
    /// <returns>Whether the cursor state was persisted.</returns>
    private static bool WriteCursor(string path, string? name)
    {
        try
        {
            SecureFile.ReplaceAtomically(path, name ?? string.Empty);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    ///     One destination a publish was working on, and what was there before it.
    /// </summary>
    /// <param name="Destination">The canonical destination.</param>
    /// <param name="Backup">
    ///     Where the previous file was moved, or null when the destination was newly created.
    /// </param>
    /// <remarks>
    ///     A step with no backup is a file this publish <em>created</em>. Recording only the ones
    ///     with backups meant a batch that wrote a hundred new images journalled nothing, so a
    ///     crash after the sixtieth left sixty files behind with no record that a transaction had
    ///     been in flight (R13-F01). Recovery removes such a file only when its content still
    ///     matches the digest below, and reports it otherwise: deleting on the strength of a
    ///     journal is the one action that gives a forger something they did not already have, and
    ///     the digest is what takes that away by naming the bytes rather than the path. The older
    ///     wording said these were only ever reported, which stopped being true when the digest
    ///     check arrived (R19-DOC01).
    /// </remarks>
    /// <param name="Created">
    ///     The SHA-256, lower-case hex, of the bytes this publish was putting at the destination.
    ///     Recorded only for a created file — one with no backup — and it is what makes removing
    ///     that file safe: a journal that merely <em>names</em> a file is a request to delete
    ///     anything, and a journal that names its contents can only ask for the file the publish
    ///     actually wrote.
    /// </param>
    public sealed record Entry(string Destination, string? Backup, string? Created = null);

    /// <summary>An unfinished publish.</summary>
    /// <param name="What">What the batch was publishing, for the log line.</param>
    /// <param name="Steps">The destinations it had started on.</param>
    /// <remarks>
    ///     Deliberately carries no allowlist. It used to record the roots the publish was confined
    ///     to, and recovery read them back to decide where a file could be moved — so the
    ///     authority for a filesystem move came out of the JSON being validated. Anyone able to
    ///     drop a <c>*.publishing.json</c> into a scanned directory could name their own roots.
    ///     The roots now come from the caller, which gets them from configuration (R13-SEC01).
    /// </remarks>
    /// <param name="State">
    ///     Whether the transaction was still replacing destinations or had put them all in place.
    ///     Absent on a record written before this existed, which is read as still publishing.
    /// </param>
    /// <param name="TransactionId">
    ///     What identifies this transaction, so a record that has already been carried out is not
    ///     carried out again. A signature says who wrote a record; it says nothing about whether
    ///     the work it describes is still outstanding (R19-REC03).
    /// </param>
    /// <param name="Signature">
    ///     The HMAC over every other field, proving this server wrote it. Checked before anything
    ///     the record names is acted on, and before the filesystem is asked about any of it.
    /// </param>
    /// <param name="Machine">The host that wrote it, so a journal from elsewhere is left alone.</param>
    /// <param name="ProcessId">The process that wrote it.</param>
    /// <param name="StartedUtc">
    ///     When that process started, which is what makes the id mean something: ids are reused,
    ///     and a start time pins the record to one process rather than to a number.
    /// </param>
    public sealed record Record(
        string What,
        IReadOnlyList<Entry> Steps,
        string? Machine = null,
        int? ProcessId = null,
        DateTimeOffset? StartedUtc = null,
        string? State = null,
        string? TransactionId = null,
        string? Signature = null);

    /// <summary>A bounded page of the persistent pending inventory.</summary>
    /// <param name="Names">Present journal file names, never paths.</param>
    /// <param name="Complete">Whether the read made a full rotation through the index.</param>
    private sealed record PendingInventory(IReadOnlyList<string> Names, bool Complete);

    /// <summary>The ledger's on-disk shape.</summary>
    /// <param name="Recovered">Transaction ids, oldest first.</param>
    /// <param name="Signature">An HMAC over the ids, so the list cannot be edited or emptied.</param>
    /// <param name="Generation">
    ///     Incremented on every write and covered by the signature, so a ledger cannot be replaced
    ///     by an older *signed* one without the substitution being visible to anyone comparing two
    ///     reads. It does not detect a rollback between two starts on its own — nothing inside a
    ///     directory can anchor that directory's own history — and that boundary is stated rather
    ///     than papered over (R20-REC09).
    /// </param>
    public sealed record Ledger(
        IReadOnlyList<string> Recovered,
        string? Signature = null,
        long Generation = 0);

    /// <summary>What one recovery did.</summary>
    /// <param name="Restored">Destinations put back from their backup.</param>
    /// <param name="Refused">Journal entries recovery declined to act on, with the reason.</param>
    /// <param name="Orphaned">
    ///     Files an interrupted publish had created that recovery would not remove, because their
    ///     content no longer matches what the record says was put there. They are left on disk and
    ///     reported.
    ///     <para>
    ///         A created file whose digest <em>does</em> match is removed, and appears in
    ///         <paramref name="Restored" />: undoing a create is deleting it, and the digest is
    ///         what makes that safe. This said such files were never deleted, which stopped being
    ///         true when the digest check was added and was not corrected (R19-DOC01).
    ///     </para>
    /// </param>
    /// <param name="Consumed">
    ///     Whether the journal itself is gone: every step reached a terminal outcome and the file
    ///     was deleted. This is the only state the ledger may record. A journal kept because
    ///     something was refused is still outstanding, and marking it executed turned the next
    ///     start's retry into a deletion (R20-REC01).
    /// </param>
    public sealed record RecoveryResult(
        IReadOnlyList<string> Restored,
        IReadOnlyList<string> Refused,
        IReadOnlyList<string> Orphaned,
        bool Consumed = false);

    /// <summary>What one start-up may spend recovering, across every root it looks in.</summary>
    /// <remarks>
    ///     The per-root and per-file caps bound one directory and one file. Discovery walks the
    ///     recovery directory and every allowlisted root, so the work a start-up did was those
    ///     caps multiplied by the number of configured roots — a number the operator chooses, on
    ///     directories anything producing output can write into (R19-RES01). This is the total.
    /// </remarks>
    public sealed class RecoveryBudget
    {
        /// <summary>The most journals one start-up will consider, across all roots.</summary>
        public int Journals { get; set; } = MaxJournalsPerScan;

        /// <summary>The most journal bytes one start-up will read before any signature verifies.</summary>
        /// <remarks>
        ///     A count alone let each of the 512 journals be four megabytes of unauthenticated
        ///     JSON, read and deserialised before its HMAC was checked — about two gigabytes of
        ///     pre-authentication work an operator's output directory could ask for (R20-RES01).
        /// </remarks>
        public long Bytes { get; set; } = 64L * 1024 * 1024;

        /// <summary>When this start-up stops recovering whatever is left.</summary>
        public DateTimeOffset Deadline { get; set; } = DateTimeOffset.UtcNow.AddSeconds(30);

        /// <summary>Whether anything is left to spend.</summary>
        public bool Exhausted => Journals <= 0 || Bytes <= 0 || DateTimeOffset.UtcNow >= Deadline;

        /// <summary>
        ///     Whether the bytes or the time are spent, regardless of how many journals are left. The
        ///     count is taken for a whole root at once, so inside that root it is already applied;
        ///     asking <see cref="Exhausted" /> there refused every journal of a root that used the
        ///     last of the count.
        /// </summary>
        public bool BytesOrTimeSpent => Bytes <= 0 || DateTimeOffset.UtcNow >= Deadline;

        /// <summary>How long is left, for a wait that has to fit inside this budget.</summary>
        public TimeSpan Remaining
        {
            get
            {
                var left = Deadline - DateTimeOffset.UtcNow;
                return left > TimeSpan.Zero ? left : TimeSpan.Zero;
            }
        }

        /// <summary>Charges the bytes one journal is about to cost.</summary>
        /// <param name="length">The file's length.</param>
        /// <returns><c>false</c> when it would exceed what is left, in which case nothing is charged.</returns>
        public bool TryCharge(long length)
        {
            if (length > Bytes) return false;
            Bytes -= length;
            return true;
        }

        /// <summary>Charges bytes that have already been read, whatever came of reading them.</summary>
        /// <param name="length">How many came in.</param>
        /// <remarks>
        ///     Unconditional: the work is done, so the budget records it, even past zero. A budget
        ///     that only recorded reads which parsed was not bounding the reads (R22-REC01).
        /// </remarks>
        public void Charge(long length)
        {
            Bytes -= Math.Max(0, length);
        }

        /// <summary>Takes what a root is about to use.</summary>
        /// <param name="count">How many journals it wants to consider.</param>
        /// <returns>How many it may.</returns>
        public int Take(int count)
        {
            var allowed = Math.Min(count, Math.Max(0, Journals));
            Journals -= allowed;
            return allowed;
        }
    }
}
