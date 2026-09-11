using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Proves that a record on disk was written by this installation, so recovery can act on it.
///     <para>
///         The cleanup queue and the publish journal are both files in a directory this server
///         scans at start-up, and both were treated as instructions: a path in them was a path to
///         delete or to restore, subject only to shape and containment rules. Containment is not
///         provenance. A record naming a file inside an allowed root is indistinguishable from one
///         this server wrote, and in the default open-world configuration there is no allowed root
///         to be inside of at all (R17-S01, R17-S02, R17-F02).
///     </para>
///     <para>
///         So each record carries an HMAC over its own content, keyed by a secret this server
///         generates once and keeps beside them. An attacker who can write into that directory can
///         still write a file; what they cannot do is make it verify. That is the difference
///         between "this path looks acceptable" and "this server asked for this".
///     </para>
///     <para>
///         <b>What the key is and is not.</b> It is created with <c>FileMode.CreateNew</c> and, on
///         Unix, mode 0600, so it belongs to the account the server runs as. It is not protection
///         against that same account — anyone who can read the file can sign anything, and the
///         threat this addresses is a <em>caller</em> who can cause writes through this server's
///         own tools, not an operator at the console. If the key cannot be created or read,
///         recovery does nothing rather than falling back to trusting the records: an
///         unauthenticated instruction is the thing being removed, so there is no safe fallback
///         to it.
///     </para>
/// </summary>
public sealed class RecoveryCapability
{
    /// <summary>The file holding this installation's signing key.</summary>
    public const string KeyFileName = ".recovery-key";

    /// <summary>How many bytes of key are generated. 256 bits, to match the HMAC.</summary>
    private const int KeyBytes = 32;

    private readonly byte[] _key;

    /// <summary>Creates a capability from a key.</summary>
    /// <param name="key">The signing key.</param>
    private RecoveryCapability(byte[] key)
    {
        _key = key;
    }

    /// <summary>
    ///     Loads the key kept in a directory, creating one if this is the first start.
    /// </summary>
    /// <param name="directory">The directory the records live in.</param>
    /// <returns>The capability, or null when no key could be established.</returns>
    /// <remarks>
    ///     Null is a refusal, not a fallback. A caller that cannot get a capability must not act on
    ///     the records it was going to check with it — the whole point is that an unverified record
    ///     is not an instruction.
    ///     <para>
    ///         An existing key is checked before it is trusted: it must be a regular file, not a
    ///         link or reparse point, and on Unix it must not be readable by group or other. It used
    ///         to be accepted on length alone, so a key planted before the server's first start was
    ///         adopted as the server's own and every forged record signed with it verified
    ///         (R18-SEC01).
    ///     </para>
    /// </remarks>
    public static RecoveryCapability? For(string directory)
    {
        var path = Path.Combine(directory, KeyFileName);

        try
        {
            // A key is only as private as the directory holding it, so a root that cannot be
            // shown to be this account's alone is a refusal rather than a place to write one.
            if (!CreateOwnerOnlyDirectory(directory)) return null;

            var existing = ReadTrustedKey(path);
            if (existing != null) return new RecoveryCapability(existing);

            // Nothing usable is there. If something *is* there and it is not usable, that is a
            // refusal: replacing it would destroy the key a running host is signing with, and
            // trusting it is what this is here to stop.
            if (File.Exists(path) || Directory.Exists(path)) return null;

            var key = RandomNumberGenerator.GetBytes(KeyBytes);

            try
            {
                WriteNewKey(path, key);
            }
            catch (IOException)
            {
                // Someone else created it in the meantime, which is the outcome that matters: both
                // servers end up using the same key — provided it passes the same checks.
                var raced = ReadTrustedKey(path);
                return raced == null ? null : new RecoveryCapability(raced);
            }

            return new RecoveryCapability(key);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or NotSupportedException or ArgumentException)
        {
            return null;
        }
    }

    /// <summary>Writes a new key, refusing an existing name and narrowing it as it is created.</summary>
    /// <param name="path">Where the key goes.</param>
    /// <param name="key">The key bytes.</param>
    /// <exception cref="IOException">Thrown when the name already exists.</exception>
    /// <remarks>
    ///     <c>CreateNew</c> so two servers starting together cannot have one overwrite the key the
    ///     other has begun signing with. On Unix the mode is set as the file is made rather than
    ///     afterwards: a key that is world-readable for an instant has been readable, and the
    ///     previous version chmod'd after closing the handle and swallowed a failure to do so.
    ///     Windows has no mode to set here and inherits the directory's ACL.
    /// </remarks>
    private static void WriteNewKey(string path, byte[] key)
    {
        var options = new FileStreamOptions
        {
            Mode = FileMode.CreateNew,
            Access = FileAccess.Write,
            Share = FileShare.None
        };

        if (!OperatingSystem.IsWindows())
            options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;

        using var stream = new FileStream(path, options);
        stream.Write(key);
    }

    /// <summary>Reads an existing key, or null when there is none this may trust.</summary>
    /// <param name="path">The key file.</param>
    /// <returns>The key bytes, or null.</returns>
    private static byte[]? ReadTrustedKey(string path)
    {
        try
        {
            // Opened first, and the questions the handle can answer asked of it: owner, length
            // and content. The previous version checked a `FileInfo` and then read through a
            // second open by path, so the file it vouched for did not have to be the file it read
            // (R19-REC01). Two questions below are still asked by name — whether it is a link and
            // whether it is a directory — because .NET exposes no handle-based attributes; that is
            // a weaker guarantee than "everything on the handle", and it is stated here rather
            // than implied (R20-REC08).
            //
            // `FileShare.Read`, not `None`. Denying writers is what keeps the file still for the
            // length of the read; denying readers as well stopped two hosts sharing a recovery
            // root from both establishing a capability — measured, one of two racing processes
            // failed every one of its thirty records for want of a key it was entitled to.
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

            // A link is a key somewhere else, chosen by whoever made the link. Asked by name: with
            // the handle held open `FileShare.Read`, the name can still be repointed, so this
            // refuses a link that was there at the open and cannot see one made after it.
            var file = new FileInfo(path);
            if (file.LinkTarget != null) return null;
            if (file.Attributes.HasFlag(FileAttributes.ReparsePoint)) return null;
            if (file.Attributes.HasFlag(FileAttributes.Directory)) return null;

            if (stream.Length != KeyBytes) return null;

            // Whose it is, which is the question that decides this. What the file looks like — its
            // length, its type — is entirely the attacker's choice; its owner is not.
            if (!BelongsToTrustedOwner(stream, path)) return null;

            var bytes = new byte[KeyBytes];
            var read = 0;
            while (read < bytes.Length)
            {
                var step = stream.Read(bytes, read, bytes.Length - read);
                if (step <= 0) return null;
                read += step;
            }

            // Nothing may follow the key. A file that is longer than it claimed between the length
            // check and the read is not the file that was measured.
            return stream.ReadByte() == -1 ? bytes : null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or NotSupportedException)
        {
            return null;
        }
    }

    /// <summary>Whether the file behind an open handle belongs to a trusted owner.</summary>
    /// <param name="stream">The open key file.</param>
    /// <param name="path">Its path, for the platforms whose API takes one.</param>
    /// <returns><c>true</c> when it does, and <c>false</c> when it does not or cannot be told.</returns>
    /// <remarks>
    ///     Windows can answer this from the handle itself, which is what makes the answer about
    ///     the file that is open rather than about whatever the name points at now. Unix has no
    ///     uid through .NET, so it falls back to the mode — sound here for the same reason as the
    ///     directory: a file nobody but its owner may read, that this process is reading, is this
    ///     process's own.
    /// </remarks>
    private static bool BelongsToTrustedOwner(FileStream stream, string path)
    {
        if (!OperatingSystem.IsWindows()) return OwnedByThisUnixUser(path);

        try
        {
            var owner = stream.GetAccessControl().GetOwner(typeof(SecurityIdentifier));
            using var identity = WindowsIdentity.GetCurrent();

            return identity.User is { } self
                   && owner is SecurityIdentifier sid && IsTrustedWindowsIdentity(sid, self);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or PrivilegeNotHeldException
                                       or PlatformNotSupportedException)
        {
            return false;
        }
    }

    /// <summary>Whether a file is owner-only, which on Unix is what this can establish.</summary>
    /// <param name="path">The file to ask about.</param>
    /// <returns><c>true</c> when nothing beyond its owner may read it.</returns>
    /// <remarks>
    ///     .NET exposes the mode but not the uid, so the test is "nobody else can read it". A file
    ///     another account owns and left at 0600 is not readable by this one, so it would fail to
    ///     load rather than be trusted.
    /// </remarks>
    [UnsupportedOSPlatform("windows")]
    private static bool OwnedByThisUnixUser(string path)
    {
        // The mode inference below — "only the owner may open it, and we opened it, so it is
        // ours" — is sound for an ordinary account and false for root, which opens anything. .NET
        // exposes no uid to compare, so the one cheap question is whether we *are* root; if so
        // nothing here can establish ownership, and the answer is no (R20-REC04).
        if (RunningAsRoot()) return false;

        try
        {
            var beyondOwner = File.GetUnixFileMode(path)
                              & ~(UnixFileMode.UserRead | UnixFileMode.UserWrite
                                                        | UnixFileMode.UserExecute);
            return beyondOwner == UnixFileMode.None;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or PlatformNotSupportedException)
        {
            return false;
        }
    }

    /// <summary>Creates the recovery directory, owner-only where the platform expresses that.</summary>
    /// <param name="directory">The directory to create.</param>
    /// <returns>
    ///     <c>true</c> when the directory exists and only this account can reach it. <c>false</c>
    ///     is a refusal to use it at all: a root that cannot be shown to be private is no place to
    ///     keep a signing key (R19-REC01).
    /// </returns>
    private static bool CreateOwnerOnlyDirectory(string directory)
    {
        try
        {
            // An existing name is not an existing private directory. This returned the moment
            // something was there, so a reparse point, a directory anyone could write to, or one
            // belonging to another account silently became this host's recovery root
            // (R19-REC01).
            if (!Directory.Exists(directory) && !File.Exists(directory))
            {
                if (OperatingSystem.IsWindows()) Directory.CreateDirectory(directory);
                else CreateUnixOwnerOnlyDirectory(directory);
            }

            // Type and ownership before anything is written to it: this must never narrow, widen
            // or otherwise touch the permissions of a directory belonging to somebody else.
            if (!IsARealDirectoryOwnedByTrustedIdentity(directory)) return false;

            // Already right is the ordinary case — every start after the first — and writing the
            // ACL anyway is not free: two hosts establishing the same root at once collided on
            // `SetAccessControl`, and the one that lost got no capability and silently recorded
            // nothing. Measured under load, one of two racing processes dropped all thirty of its
            // debts. Narrow it only when it needs narrowing.
            // Twice, because two processes starting together can both find it not yet private
            // and then collide: the loser sees a directory another process is part-way through
            // narrowing, and one look is not enough to tell that from a directory that is wrong.
            for (var attempt = 0; attempt < 2; attempt++)
            {
                if (IsAPrivateDirectoryOfThisAccount(directory)) return true;

                MakePrivate(directory);
            }

            return IsAPrivateDirectoryOfThisAccount(directory);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>Whether a path is a real directory, not a link, with a trusted owner.</summary>
    /// <param name="directory">The directory to judge.</param>
    /// <returns><c>true</c> when it is; <c>false</c> when it is not or cannot be told.</returns>
    private static bool IsARealDirectoryOwnedByTrustedIdentity(string directory)
    {
        try
        {
            var info = new DirectoryInfo(directory);
            if (!info.Exists) return false;

            // A link is a directory somewhere else, chosen by whoever made the link.
            if (info.LinkTarget != null) return false;
            if (info.Attributes.HasFlag(FileAttributes.ReparsePoint)) return false;

            if (!OperatingSystem.IsWindows()) return true;

            using var identity = WindowsIdentity.GetCurrent();
            return identity.User is { } self
                   && info.GetAccessControl().GetOwner(typeof(SecurityIdentifier))
                       is SecurityIdentifier owner && IsTrustedWindowsIdentity(owner, self);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or PrivilegeNotHeldException
                                       or PlatformNotSupportedException)
        {
            return false;
        }
    }

    /// <summary>Narrows a directory this account owns so only this account can reach it.</summary>
    /// <param name="directory">The directory to narrow.</param>
    /// <remarks>
    ///     Set, not assumed. A directory created under the user's temp root inherits whatever that
    ///     root grants, which on an ordinary machine is several identities that are not the owner;
    ///     verifying the inherited ACL would refuse every such machine, and trusting it would
    ///     accept a root other accounts can write into. Failures are left to the verification that
    ///     follows, which refuses rather than proceeds.
    /// </remarks>
    private static void MakePrivate(string directory)
    {
        try
        {
            if (OperatingSystem.IsWindows()) MakeWindowsPrivate(directory);
            else MakeUnixPrivate(directory);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or PrivilegeNotHeldException
                                       or PlatformNotSupportedException)
        {
            // Reported by the check that follows, which is the one that decides.
        }
    }

    /// <summary>Replaces a directory's ACL with one that inherits nothing and grants three.</summary>
    /// <param name="directory">The directory to narrow.</param>
    [SupportedOSPlatform("windows")]
    private static void MakeWindowsPrivate(string directory)
    {
        using var identity = WindowsIdentity.GetCurrent();
        if (identity.User is not { } self) return;

        var info = new DirectoryInfo(directory);
        var security = info.GetAccessControl();

        // Inheritance off and the inherited rules dropped rather than copied: copying them is what
        // would carry the temp root's grants into the private directory.
        security.SetAccessRuleProtection(true, false);

        foreach (FileSystemAccessRule existing in
                 security.GetAccessRules(true, false, typeof(SecurityIdentifier)))
            security.RemoveAccessRuleSpecific(existing);

        foreach (var sid in TrustedSids(self))
            security.AddAccessRule(new FileSystemAccessRule(sid, FileSystemRights.FullControl,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                PropagationFlags.None, AccessControlType.Allow));

        info.SetAccessControl(security);
    }

    /// <summary>The only identities the recovery root grants anything to.</summary>
    /// <param name="self">This process's user.</param>
    /// <returns>This account, the system account and the local administrators.</returns>
    /// <remarks>
    ///     The last two can read anything on the machine whatever this ACL says, so excluding them
    ///     would deny nobody anything while making the directory unmanageable.
    /// </remarks>
    [SupportedOSPlatform("windows")]
    private static IEnumerable<SecurityIdentifier> TrustedSids(SecurityIdentifier self)
    {
        yield return self;
        yield return new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
        yield return new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null);
    }

    /// <summary>Narrows a Unix directory's mode to its owner.</summary>
    /// <param name="directory">The directory to narrow.</param>
    [UnsupportedOSPlatform("windows")]
    private static void MakeUnixPrivate(string directory)
    {
        File.SetUnixFileMode(directory,
            UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
    }

    /// <summary>Whether a path is a real directory that only this account can reach.</summary>
    /// <param name="directory">The directory to judge.</param>
    /// <returns><c>true</c> when it is; <c>false</c> when it is not or cannot be told.</returns>
    /// <remarks>
    ///     Fail closed. Unknown answers "no", because the whole point of this directory is that
    ///     nobody else can put a key or a journal in it, and a directory that cannot be shown to
    ///     be private is one there is no reason to keep secrets in (R19-REC01).
    /// </remarks>
    private static bool IsAPrivateDirectoryOfThisAccount(string directory)
    {
        try
        {
            if (!IsARealDirectoryOwnedByTrustedIdentity(directory)) return false;

            return OperatingSystem.IsWindows()
                ? IsAPrivateWindowsDirectory(directory)
                : IsAPrivateUnixDirectory(directory);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>Whether a Windows directory is owned by this account and shared with nobody.</summary>
    /// <param name="directory">The directory.</param>
    /// <returns><c>true</c> when the owner is this account and no other identity is granted access.</returns>
    /// <remarks>
    ///     Owner alone is not enough: a directory this account owns can still carry an inherited
    ///     ACL that lets another user write into it, and the previous version created it with
    ///     exactly that inherited ACL and never looked.
    /// </remarks>
    [SupportedOSPlatform("windows")]
    private static bool IsAPrivateWindowsDirectory(string directory)
    {
        try
        {
            var security = new DirectoryInfo(directory).GetAccessControl();
            using var identity = WindowsIdentity.GetCurrent();

            if (identity.User is not { } self) return false;

            // Nothing inherited: the protection flag is what stops the temp root's grants applying
            // here, and a root that still inherits is one whose contents someone else may reach.
            if (!security.AreAccessRulesProtected) return false;

            foreach (FileSystemAccessRule rule in
                     security.GetAccessRules(true, true, typeof(SecurityIdentifier)))
            {
                if (rule.AccessControlType != AccessControlType.Allow) continue;
                if (rule.IdentityReference is not SecurityIdentifier granted) return false;
                if (!IsTrustedWindowsIdentity(granted, self)) return false;
            }

            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or PrivilegeNotHeldException
                                       or PlatformNotSupportedException)
        {
            return false;
        }
    }

    /// <summary>Whether an identity granted access to the recovery root is one this trusts.</summary>
    /// <param name="granted">The identity in the access rule.</param>
    /// <param name="self">This process's user.</param>
    /// <returns><c>true</c> when the grant does not widen who can reach the directory.</returns>
    /// <remarks>
    ///     This account, plus the two the operating system puts on everything it creates. An
    ///     administrator and the system account can read anything on the machine regardless of
    ///     this ACL, so refusing their presence would refuse every directory Windows makes without
    ///     denying anyone anything.
    /// </remarks>
    [SupportedOSPlatform("windows")]
    private static bool IsTrustedWindowsIdentity(SecurityIdentifier granted,
        SecurityIdentifier self)
    {
        if (granted.Equals(self)) return true;
        if (granted.IsWellKnown(WellKnownSidType.LocalSystemSid)) return true;
        if (granted.IsWellKnown(WellKnownSidType.BuiltinAdministratorsSid)) return true;

        return false;
    }

    /// <summary>Whether a Unix directory grants nothing beyond its owner.</summary>
    /// <param name="directory">The directory.</param>
    /// <returns><c>true</c> when its mode is owner-only.</returns>
    /// <remarks>
    ///     .NET exposes the mode but not the uid, so ownership is inferred rather than read: a
    ///     directory whose mode grants nothing to group or other can only be entered by its owner,
    ///     so a process that can enter it is the owner (or root). Enforcing the mode on an
    ///     <em>existing</em> directory is what was missing — it was only ever set on one this
    ///     server created.
    /// </remarks>
    [UnsupportedOSPlatform("windows")]
    private static bool IsAPrivateUnixDirectory(string directory)
    {
        // The mode inference below — "only the owner may open it, and we opened it, so it is
        // ours" — is sound for an ordinary account and false for root, which opens anything. .NET
        // exposes no uid to compare, so the one cheap question is whether we *are* root; if so
        // nothing here can establish ownership, and the answer is no (R20-REC04).
        if (RunningAsRoot()) return false;

        try
        {
            var beyondOwner = File.GetUnixFileMode(directory)
                              & ~(UnixFileMode.UserRead | UnixFileMode.UserWrite
                                                        | UnixFileMode.UserExecute);
            return beyondOwner == UnixFileMode.None;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or PlatformNotSupportedException)
        {
            return false;
        }
    }

    /// <summary>Creates a directory only its owner may enter.</summary>
    /// <param name="directory">The directory to create.</param>
    [UnsupportedOSPlatform("windows")]
    private static void CreateUnixOwnerOnlyDirectory(string directory)
    {
        Directory.CreateDirectory(directory,
            UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
    }

    /// <summary>Signs a record's content.</summary>
    /// <param name="content">The exact text that will be stored.</param>
    /// <returns>The signature, lower-case hex.</returns>
    public string Sign(string content)
    {
        return Convert.ToHexString(HMACSHA256.HashData(_key, Encoding.UTF8.GetBytes(content)))
            .ToLowerInvariant();
    }

    /// <summary>Whether a signature is one this installation produced for this content.</summary>
    /// <param name="content">The stored text.</param>
    /// <param name="signature">The signature stored with it.</param>
    /// <returns><c>true</c> when the record verifies.</returns>
    /// <remarks>
    ///     Compared in fixed time. A byte-by-byte comparison that returns early tells an attacker
    ///     how much of a guess was right, which is how a forgery is built one character at a time.
    /// </remarks>
    public bool Verify(string content, string? signature)
    {
        if (string.IsNullOrEmpty(signature)) return false;

        var expected = Encoding.ASCII.GetBytes(Sign(content));
        var actual = Encoding.ASCII.GetBytes(signature);

        return CryptographicOperations.FixedTimeEquals(expected, actual);
    }

    /// <summary>Whether this process's effective user is root.</summary>
    /// <returns><c>true</c> when it is, or when that cannot be determined.</returns>
    /// <remarks>
    ///     Unknown answers "yes", which refuses: a platform where <c>geteuid</c> cannot be called
    ///     is one where the mode inference cannot be trusted either. <c>uid_t</c> is a 32-bit
    ///     unsigned integer on every libc this server runs on, which is what makes this the one
    ///     ownership question that can be asked without a platform-specific struct layout.
    /// </remarks>
    [UnsupportedOSPlatform("windows")]
    private static bool RunningAsRoot()
    {
        try
        {
            return geteuid() == 0;
        }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException)
        {
            return true;
        }
    }

    [SuppressMessage("Interoperability", "SYSLIB1054",
        Justification =
            "DllImport preferred over LibraryImport to avoid AllowUnsafeBlocks; these APIs are called infrequently")]
    [DllImport("libc", EntryPoint = "geteuid")]
    private static extern uint geteuid();
}
