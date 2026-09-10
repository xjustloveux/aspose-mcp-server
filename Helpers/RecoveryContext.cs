using System.Collections.Concurrent;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Where one host keeps its recovery records, and the key that proves it wrote them.
///     <para>
///         These were two mutable statics — <c>PublishJournal.JournalDirectory</c> and a cached
///         capability beside it — rewritten by whichever <c>CleanupDebtService</c> started last.
///         Two hosts in one process is a shape the debt-sink stack was built for, and under it the
///         second to start moved the first's journals, signed its queue with a different key, and
///         left records the first would refuse at its next start (R18-ARCH01).
///     </para>
///     <para>
///         Immutable, and obtained per directory rather than per process. Anything that writes a
///         record is handed the context to write it into, so there is no ambient answer to "which
///         host is this" for the wrong one to supply.
///     </para>
///     <para>
///         The records live in a subdirectory this class creates, not in the directory it is given.
///         A bare system temp root is a shared namespace, and a key sitting in one can be replaced
///         before the server first reads it (R18-SEC01); <see cref="RecoveryCapability" /> checks
///         what it finds, and putting it somewhere the server made is the first half of that.
///     </para>
/// </summary>
public sealed class RecoveryContext
{
    /// <summary>The subdirectory recovery records are kept in.</summary>
    public const string DirectoryName = ".aspose-recovery";

    /// <summary>Contexts already established, one per canonical directory.</summary>
    /// <remarks>
    ///     A cache, not shared mutable state: each entry is immutable once created and keyed by the
    ///     directory it belongs to. Two hosts on two directories get two contexts; two components
    ///     on one directory get the same one, which is what lets each verify the other's records.
    /// </remarks>
    private static readonly ConcurrentDictionary<string, RecoveryContext> Established =
        new(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);

    /// <summary>Creates a context.</summary>
    /// <param name="directory">Where the records live.</param>
    /// <param name="capability">The key they are signed with, or null when none could be made.</param>
    private RecoveryContext(string directory, RecoveryCapability? capability)
    {
        Directory = directory;
        Capability = capability;
    }

    /// <summary>Where this host's journals and queue live, and the only place it recovers from.</summary>
    public string Directory { get; }

    /// <summary>
    ///     The key this host signs and verifies with, or null when none could be established.
    /// </summary>
    /// <remarks>
    ///     Null is a refusal, not a fallback. A component that cannot sign must not write a record
    ///     recovery would ignore, and one that cannot verify must not act on a record at all.
    /// </remarks>
    public RecoveryCapability? Capability { get; }

    /// <summary>The context for a host's temp directory, establishing its key on first use.</summary>
    /// <param name="temporaryDirectory">The host's temp directory.</param>
    /// <returns>The context.</returns>
    public static RecoveryContext For(string temporaryDirectory)
    {
        var key = Canonical(Path.Combine(temporaryDirectory, DirectoryName));

        var context = Established.GetOrAdd(key,
            resolved => new RecoveryContext(resolved, RecoveryCapability.For(resolved)));

        // A context with no capability is a refusal, and a refusal is not worth remembering: the
        // cause may be a lock another process held for a moment or a directory not yet writable,
        // and caching it made one transient failure a permanent "no key" for the life of the
        // process (R20-REC07). Dropped so the next caller tries again; a genuine, lasting refusal
        // is simply re-established each time and costs a directory check.
        if (context.Capability == null) Established.TryRemove(key, out _);

        return context;
    }

    /// <summary>Canonicalises a directory, tolerating one that cannot be resolved.</summary>
    /// <param name="directory">The directory as configured.</param>
    /// <returns>Its canonical form, or the original when it has none.</returns>
    private static string Canonical(string directory)
    {
        try
        {
            return Path.GetFullPath(directory);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException
                                       or PathTooLongException or IOException)
        {
            return directory;
        }
    }
}
