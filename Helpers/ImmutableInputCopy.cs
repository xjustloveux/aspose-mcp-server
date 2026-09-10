namespace AsposeMcpServer.Helpers;

/// <summary>
///     A private copy of an input, so a security decision and the work it authorises are made
///     about the same bytes.
///     <para>
///         Conversion decided whether an input was safe by opening the caller's path, and then
///         handed that same path to the parser, which opened it again. Between the two opens the
///         file is the caller's to replace — and for the formats this guards, the scan is the only
///         control there is, so a swap converts an unscanned document (R19-CNV01).
///     </para>
///     <para>
///         The copy lives under the host's own recovery root, and only when that root has been
///         shown to be private: a host with no <see cref="RecoveryContext.Capability" /> is one
///         whose root could not be, and copying a caller's whole input there anyway made the
///         privacy failure a data sink (R20-CNV01). Its bucket carries a nonce and is claimed
///         atomically, so nothing can be waiting at the copy path.
///     </para>
///     <para>
///         The copy is removed when the conversion finishes. One left behind by a crash is inert —
///         nonce-bucketed, in a directory only this account can reach — and is removed by the bounded
///         sweep <c>CleanupDebtService</c> runs over this subdirectory at start-up (R20-CNV03).
///         The first four nonce characters select one of 65,536 two-level buckets. A bucket is
///         claimed with <see cref="FileMode.CreateNew" /> and holds at most one copy, so the
///         sweep can rotate bounded probes without replaying a non-deletable directory prefix.
///     </para>
///     <para>
///         For as long as this object lives it holds an exclusive handle on a lease file beside
///         the copy. The copy itself has no handle on it between the scan and the loader
///         and between one loader open and the next, and a sweep in another host that probed the
///         copy in one of those instants found nothing holding it and removed a conversion's
///         input from under it (R22-REC02). The lease is what the sweep must take before it
///         deletes, and it cannot take one a running conversion holds. Delete-on-close, so a
///         crash releases it with the process and leaves the copy plainly abandoned.
///     </para>
///     <para>
///         Weaker on Unix: <see cref="FileShare.None" /> is advisory there, honoured between .NET
///         processes and not enforced by the kernel against others.
///     </para>
/// </summary>
public sealed class ImmutableInputCopy : IDisposable
{
    /// <summary>The subdirectory copies are kept in, under a host's recovery root.</summary>
    public const string DirectoryName = "authorised-inputs";

    /// <summary>What follows a copy's name to name its lease.</summary>
    public const string LeaseSuffix = ".lease";

    private const string ClaimFileName = ".claim";
    private readonly string? _claim;
    private readonly FileStream _lease;

    /// <summary>Creates the handle.</summary>
    /// <param name="path">Where the copy is.</param>
    /// <param name="lease">The exclusive handle on the copy's lease, held until disposal.</param>
    /// <param name="claim">The bucket claim which prevents a second copy entering this bucket.</param>
    private ImmutableInputCopy(string path, FileStream lease, string? claim = null)
    {
        Path = path;
        _lease = lease;
        _claim = claim;
    }

    /// <summary>The copy every reader of this input must open.</summary>
    public string Path { get; }

    /// <summary>
    ///     Called once the copy exists, before anything reads it.
    ///     <para>
    ///         A seam, because the race this class closes cannot be reached from a fixture
    ///         otherwise: a test needs to replace the caller's file at exactly the moment between
    ///         the copy being taken and the scan being run, and no amount of threading makes that
    ///         reliable. Firing here is that moment, exactly.
    ///     </para>
    /// </summary>
    internal static Action<string>? AfterCopy { get; set; }

    /// <summary>
    ///     Opens the source, for a fixture that needs a source which grows while it is copied.
    ///     Production opens a <see cref="FileStream" /> with <see cref="FileShare.Read" />.
    /// </summary>
    internal static Func<string, Stream>? OpenSource { get; set; }

    /// <summary>Removes the copy, then releases its lease.</summary>
    /// <remarks>
    ///     In that order: the copy goes while the lease is still held, so no sweep can be
    ///     deleting it at the same time; the lease then deletes itself on close.
    /// </remarks>
    public void Dispose()
    {
        try
        {
            if (File.Exists(Path)) File.Delete(Path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A copy left behind is inert: its claimed bucket carries a nonce, so nothing will
            // open it again, and the start-up sweep removes it once the lease is gone.
        }
        finally
        {
            _lease.Dispose();
            ReleaseBucket(_claim);
        }
    }

    /// <summary>The lease file that belongs to a copy.</summary>
    /// <param name="copy">The copy's path.</param>
    /// <returns>The path of its lease.</returns>
    public static string LeasePathOf(string copy)
    {
        var directory = System.IO.Path.GetDirectoryName(copy);
        return directory != null && File.Exists(System.IO.Path.Combine(directory, ClaimFileName))
            ? System.IO.Path.Combine(directory, LeaseSuffix)
            : copy + LeaseSuffix;
    }

    /// <summary>
    ///     Opens a copy's lease exclusively, delete-on-close.
    /// </summary>
    /// <param name="copy">The copy's path.</param>
    /// <param name="mode">
    ///     <see cref="FileMode.CreateNew" /> for the conversion that owns the copy;
    ///     <see cref="FileMode.OpenOrCreate" /> for a sweep, which must hold it while it deletes.
    /// </param>
    /// <returns>The handle; while it is open nobody else can open the lease.</returns>
    /// <exception cref="IOException">Thrown when another process holds the lease.</exception>
    public static FileStream OpenLease(string copy, FileMode mode)
    {
        return new FileStream(LeasePathOf(copy), mode, FileAccess.ReadWrite, FileShare.None, 1,
            FileOptions.DeleteOnClose);
    }

    /// <summary>Copies an input somewhere only this server can reach.</summary>
    /// <param name="source">The caller's path, already resolved and allowlisted at the entry.</param>
    /// <param name="recovery">The host whose private root the copy goes under.</param>
    /// <param name="allowedBasePaths">
    ///     The allowlist the source is resolved against again, immediately before it is opened.
    /// </param>
    /// <param name="maxBytes">
    ///     The most the input may be. Decided on the opened source's length before a byte is
    ///     copied, and enforced again on the bytes that actually come in, so a source that grows
    ///     under the copy is refused at the cap rather than staged whole (R23-RES01).
    /// </param>
    /// <returns>A handle to the copy, which the caller disposes.</returns>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the host has no recovery capability: its root could not be shown to be
    ///     private, so there is nowhere this copy may go.
    /// </exception>
    /// <exception cref="ArgumentException">
    ///     Thrown when the source resolves outside the allowlist, or is larger than
    ///     <paramref name="maxBytes" />.
    /// </exception>
    /// <exception cref="IOException">Thrown when the input could not be copied.</exception>
    /// <remarks>
    ///     The source is opened <em>once</em>, with <see cref="FileShare.Read" /> so no writer can
    ///     change it while it is being copied, and the bytes come from that handle rather than
    ///     from a second path-based read (R20-CNV02). What remains is the window between the
    ///     resolution and the open — the same one <c>R18-SEC03</c> could not close — which is why
    ///     the resolution happens here, immediately before the open, and not only at the entry.
    /// </remarks>
    public static ImmutableInputCopy Of(string source, RecoveryContext recovery,
        IReadOnlyList<string> allowedBasePaths, long maxBytes = long.MaxValue)
    {
        if (recovery.Capability == null)
            throw new InvalidOperationException(
                "This host has no recovery signing key, so its recovery root could not be shown "
                + "to be private and no input may be staged there.");

        var resolvedSource =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(source, allowedBasePaths, nameof(source));

        var extension = System.IO.Path.GetExtension(resolvedSource);
        var staging = System.IO.Path.Combine(recovery.Directory, DirectoryName);
        var (directory, claim) = ClaimBucket(staging);
        var copy = System.IO.Path.Combine(directory, "input" + extension);

        // The lease before the copy: from the first byte written there is a holder a sweep
        // cannot get past (R22-REC02).
        var lease = OpenLease(copy, FileMode.CreateNew);
        try
        {
            using var from = OpenSource?.Invoke(resolvedSource)
                             ?? new FileStream(resolvedSource, FileMode.Open, FileAccess.Read, FileShare.Read);
            // Admission on the handle being copied, before a byte is staged: the scanners'
            // size limits used to be applied to the copy, after the whole input had been
            // written into the staging area (R23-RES01).
            if (from.CanSeek && from.Length > maxBytes)
                throw new ArgumentException(
                    $"The input is {from.Length:N0} bytes, above the {maxBytes:N0} byte limit.",
                    nameof(source));

            using var to = new FileStream(copy, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            var buffer = new byte[81920];
            long copied = 0;
            int read;
            while ((read = from.Read(buffer, 0, buffer.Length)) > 0)
            {
                copied += read;
                // Enforced on what arrives, not on what the length promised: a source that
                // grows under the copy stops here, at the cap, with the lease still held.
                if (copied > maxBytes)
                    throw new ArgumentException(
                        $"The input grew past the {maxBytes:N0} byte limit while it was being staged.",
                        nameof(source));
                to.Write(buffer, 0, read);
            }
        }
        catch
        {
            // Under the lease, so no sweep can be reaching for the partial copy at the same
            // moment; then the lease deletes itself.
            try
            {
                if (File.Exists(copy)) File.Delete(copy);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Nonce-named and inert; the start-up sweep removes it once the lease is gone.
            }

            lease.Dispose();
            ReleaseBucket(claim);
            throw;
        }

        AfterCopy?.Invoke(resolvedSource);

        return new ImmutableInputCopy(copy, lease, claim);
    }

    /// <summary>Claims one randomly selected two-level bucket for exactly one staged copy.</summary>
    /// <param name="staging">The private staged-input root.</param>
    /// <returns>The bucket and its claim file.</returns>
    /// <exception cref="IOException">Thrown when no free bucket is found after bounded retries.</exception>
    private static (string Directory, string Claim) ClaimBucket(string staging)
    {
        for (var attempt = 0; attempt < StagingBudget.MaxBucketClaimAttempts; attempt++)
        {
            var nonce = Guid.NewGuid().ToString("N");
            var directory = System.IO.Path.Combine(staging, nonce[..2], nonce.Substring(2, 2));
            Directory.CreateDirectory(directory);
            var claim = System.IO.Path.Combine(directory, ClaimFileName);
            try
            {
                using (new FileStream(claim, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    // Creating the empty claim atomically reserves this bucket for one process.
                }

                return (directory, claim);
            }
            catch (IOException)
            {
                // Another process owns this bucket. A fresh random bucket avoids sharing a
                // directory, which is the invariant the bounded start-up sweep depends on.
            }
        }

        throw new IOException($"A free staged-input bucket could not be claimed after "
                              + $"{StagingBudget.MaxBucketClaimAttempts} attempts.");
    }

    /// <summary>Releases a bucket claim and removes its empty private directories when possible.</summary>
    /// <param name="claim">The claim file, or null for a legacy copy.</param>
    internal static void ReleaseBucket(string? claim)
    {
        if (claim == null) return;
        try
        {
            if (File.Exists(claim)) File.Delete(claim);
            var bucket = System.IO.Path.GetDirectoryName(claim);
            if (bucket != null && Directory.Exists(bucket)) Directory.Delete(bucket);
            var shard = bucket == null ? null : System.IO.Path.GetDirectoryName(bucket);
            if (shard != null && Directory.Exists(shard)
                              && !Directory.EnumerateFileSystemEntries(shard).Any())
                Directory.Delete(shard);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A competing creator or cleanup owns an entry; the bounded sweep can retry later.
        }
    }

    /// <summary>Releases the claimed bucket containing a copy; legacy copies have no claim.</summary>
    /// <param name="copy">The staged copy path.</param>
    internal static void ReleaseBucketFor(string copy)
    {
        var directory = System.IO.Path.GetDirectoryName(copy);
        if (directory == null) return;
        var claim = System.IO.Path.Combine(directory, ClaimFileName);
        if (File.Exists(claim)) ReleaseBucket(claim);
    }

    /// <summary>Bounds filesystem allocations while selecting a collision-free bucket.</summary>
    private static class StagingBudget
    {
        /// <summary>The most random bucket claims attempted for one staged input.</summary>
        internal const int MaxBucketClaimAttempts = 256;
    }
}
