using System.Runtime.Versioning;
using System.Security.AccessControl;
using System.Security.Principal;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R18-SEC01 and R18-ARCH01: whose key this is, and whose records it signs.
///     <para>
///         The signature that makes a journal or a queue entry an instruction is only worth the key
///         behind it. That key was accepted on length alone — thirty-two bytes in the right place,
///         no check that the file was a regular file rather than a link, no owner or mode check —
///         and it was created with the platform's default permissions and narrowed afterwards. A
///         key planted before the server's first start was adopted as the server's own, and every
///         forged record signed with it verified (R18-SEC01).
///     </para>
///     <para>
///         And there was one key and one directory for the whole process, rewritten by whichever
///         host started last. Two hosts in one process — the shape the debt-sink stack exists for —
///         then wrote into each other's directory and signed with each other's key (R18-ARCH01).
///     </para>
/// </summary>
public class RecoveryContextTests : TestBase
{
    /// <summary>A directory of this test's own to act as one host's temp directory.</summary>
    /// <param name="name">Which host.</param>
    /// <returns>The path.</returns>
    private string AHost(string name)
    {
        var path = Path.Combine(TestDir, name);
        Directory.CreateDirectory(path);
        return path;
    }

    [Fact]
    public void TwoHostsOnDifferentDirectories_ShouldNotShareADirectoryOrAKey()
    {
        // The whole of R18-ARCH01 in one assertion pair. With a process-wide directory and key,
        // the second of these to be created moved the first's journals and re-signed its queue.
        var first = RecoveryContext.For(AHost("host_one"));
        var second = RecoveryContext.For(AHost("host_two"));

        Assert.NotEqual(first.Directory, second.Directory);
        Assert.NotNull(first.Capability);
        Assert.NotNull(second.Capability);

        const string record = "a record one of them wrote";
        Assert.False(second.Capability!.Verify(record, first.Capability!.Sign(record)),
            "one host's key verified the other's record");
    }

    [Fact]
    public void TheSameDirectoryAskedTwice_ShouldGiveTheSameKey()
    {
        // The other half: components of one host must verify each other's records, or the queue
        // would refuse what the journal wrote.
        var host = AHost("host_stable");

        var first = RecoveryContext.For(host);
        var second = RecoveryContext.For(host);

        const string record = "a record this host wrote";
        Assert.True(second.Capability!.Verify(record, first.Capability!.Sign(record)));
    }

    [Fact]
    public void RecordsAreNotKeptInTheBareTempRoot()
    {
        // A bare temp root is a shared namespace. The records go in a subdirectory the server
        // makes, which is what gives the key an ownership story at all (R18-SEC01).
        var host = AHost("host_subdirectory");

        var recovery = RecoveryContext.For(host);

        Assert.NotEqual(Path.GetFullPath(host), recovery.Directory);
        Assert.Equal(RecoveryContext.DirectoryName, Path.GetFileName(recovery.Directory));
    }

    [Fact]
    public void AKeyOfTheWrongShape_ShouldNotBeAdopted()
    {
        // Length and type are the attacker's choice, so they are checked — but they are not what
        // decides this. See the ownership case below for the part that does.
        var host = AHost("host_wrong_shape");
        var directory = Path.Combine(host, RecoveryContext.DirectoryName);
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, RecoveryCapability.KeyFileName), new byte[16]);

        Assert.Null(RecoveryCapability.For(directory));
    }

    [Fact]
    public void AKeyOwnedByThisAccount_IsAdopted_AndThatIsTheBoundary()
    {
        // Written down rather than asserted away. A key planted by the account the server runs as
        // is indistinguishable from one the server wrote — nothing about a file separates them —
        // and that account can read the real key in any case, so there is nothing to defend.
        //
        // What the owner check buys is the case §30.3's attack path is about: a key placed by a
        // *different* local account in a shared temp namespace. That cannot be posed from a
        // single-account test host, so it is not claimed as verified here; what is verified is that
        // ownership is consulted at all, by the two cases that can be posed — a link, and a shape
        // this never writes.
        var host = AHost("host_same_account");
        var directory = Path.Combine(host, RecoveryContext.DirectoryName);
        Directory.CreateDirectory(directory);

        var planted = new byte[32];
        for (var i = 0; i < planted.Length; i++) planted[i] = (byte)i;
        var keyPath = Path.Combine(directory, RecoveryCapability.KeyFileName);
        File.WriteAllBytes(keyPath, planted);
        // Planted the way the server writes it: owner-only. A key anyone may read is refused on
        // Unix by the documented boundary, and the default umask leaves it group/other-readable;
        // that refusal is a different fixture's subject, not this one's.
        if (!OperatingSystem.IsWindows()) MakeOwnerOnly(keyPath);

        var capability = RecoveryCapability.For(directory);

        Assert.NotNull(capability);
        Assert.True(capability.Verify("a record", SigningWith(planted)),
            "a key this account owns is used, which is the documented boundary");
    }

    /// <summary>Narrows a file's mode to its owner.</summary>
    /// <param name="path">The file.</param>
    [UnsupportedOSPlatform("windows")]
    private static void MakeOwnerOnly(string path)
    {
        File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);
    }

    [SkippableFact]
    public void AKeyThatIsALink_ShouldNotBeRead()
    {
        // A link is a key somewhere else, chosen by whoever made the link — including somewhere the
        // attacker can read. This is one of the two ownership-adjacent cases a single-account host
        // can actually pose.
        var host = AHost("host_linked");
        var directory = Path.Combine(host, RecoveryContext.DirectoryName);
        Directory.CreateDirectory(directory);

        var elsewhere = Path.Combine(TestDir, "their_key_directory");
        Directory.CreateDirectory(elsewhere);

        Skip.IfNot(
            MidChainLinkFixture.TryCreateDirectoryLink(
                Path.Combine(directory, RecoveryCapability.KeyFileName), elsewhere),
            "This host does not allow creating a directory link");

        Assert.Null(RecoveryCapability.For(directory));
    }

    /// <summary>The signature a given key would produce, for comparing against.</summary>
    /// <param name="key">The key bytes.</param>
    /// <returns>The signature over the fixture's record.</returns>
    private string SigningWith(byte[] key)
    {
        var directory = Path.Combine(TestDir, "their_installation", RecoveryContext.DirectoryName);
        Directory.CreateDirectory(directory);
        var keyPath = Path.Combine(directory, RecoveryCapability.KeyFileName);
        File.WriteAllBytes(keyPath, key);
        if (!OperatingSystem.IsWindows()) MakeOwnerOnly(keyPath);

        var theirs = RecoveryCapability.For(directory);
        Assert.NotNull(theirs);

        return theirs.Sign("a record");
    }

    [SkippableFact]
    public void ARecoveryRootThatIsAlreadyALink_ShouldNotBeAdopted()
    {
        // R19-REC01. `CreateOwnerOnlyDirectory` returned the moment something existed at the
        // name, so a junction planted before the first start made every journal and the signing
        // key land wherever it pointed.
        var host = AHost("linked-root");
        var elsewhere = Directory.CreateDirectory(Path.Combine(TestDir, "elsewhere")).FullName;

        var root = Path.Combine(host, RecoveryContext.DirectoryName);
        Skip.IfNot(MidChainLinkFixture.TryCreateDirectoryLink(root, elsewhere),
            "This machine cannot create a directory link.");

        Assert.Null(RecoveryContext.For(host).Capability);
        Assert.Empty(Directory.GetFiles(elsewhere));
    }

    [SkippableFact]
    public void ARecoveryRootAnyoneCanWriteTo_ShouldBeNarrowedAndAdopted()
    {
        // The Unix half of `ARecoveryRootWidenedAfterwards_ShouldBeNarrowedAgain`: a directory this
        // account owns whose mode lets others in is put back to owner-only on every start, not
        // refused. This fixture used to assert refusal, contradicting the Windows design, and had
        // never run on Unix; the first container run showed the product narrowing and adopting,
        // which is what its own comments say it does (R22-DEP02 evidence run).
        Skip.If(OperatingSystem.IsWindows(),
            "File modes are the Unix expression of this; Windows is covered by the ACL fixture.");

        if (!OperatingSystem.IsWindows()) WidenOnUnix(AHost("permissive-root"));
    }

    /// <summary>Pre-creates a recovery root anyone may write to, then asks for its key.</summary>
    /// <param name="host">The host directory.</param>
    /// <remarks>Split out so the Unix-only call sits behind an attribute the analyser accepts.</remarks>
    [UnsupportedOSPlatform("windows")]
    private static void WidenOnUnix(string host)
    {
        var root = Directory
            .CreateDirectory(Path.Combine(host, RecoveryContext.DirectoryName))
            .FullName;

        File.SetUnixFileMode(root,
            UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute
            | UnixFileMode.GroupRead | UnixFileMode.GroupWrite | UnixFileMode.GroupExecute
            | UnixFileMode.OtherRead | UnixFileMode.OtherWrite | UnixFileMode.OtherExecute);

        Assert.NotNull(RecoveryContext.For(host).Capability);
        Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
            File.GetUnixFileMode(root));
    }

    [Fact]
    public void ARecoveryRootThisAccountMade_IsAdopted()
    {
        // The control. Refusing an existing directory outright would refuse every restart, so the
        // check has to accept the one this server made last time.
        var host = AHost("ordinary-root");

        var first = RecoveryContext.For(host).Capability;
        Assert.NotNull(first);

        Assert.True(Directory.Exists(Path.Combine(host, RecoveryContext.DirectoryName)));
    }

    [Fact]
    public void AKeyLongerThanAKey_ShouldNotBeRead()
    {
        // The bound, pinned on the handle-bound read. This is deliberately *not* claimed as
        // evidence for R19-REC02: measured against the previous version it passes there too,
        // because that version refused a long file through `FileInfo.Length`. What R19-REC02
        // fixes is that the length and the bytes came from two different opens, and showing that
        // needs the file swapped between them — the same race R18-SEC03 could not be made
        // deterministic either. The fix is structural; this fixture guards the boundary around it.
        // `RecoveryCapability.For` is used directly because `RecoveryContext` caches per
        // directory and would hand back the capability established before the change.
        var root = Directory
            .CreateDirectory(Path.Combine(AHost("grown-key"), RecoveryContext.DirectoryName))
            .FullName;

        Assert.NotNull(RecoveryCapability.For(root));

        var keyFile = Directory.GetFiles(root).Single();
        using (var stream = new FileStream(keyFile, FileMode.Append, FileAccess.Write))
        {
            stream.Write(new byte[8]);
        }

        Assert.Null(RecoveryCapability.For(root));
    }

    [Fact]
    public void AKeyOfExactlyTheRightLength_IsStillRead()
    {
        // The control for the bound above: tightening the read must not refuse the key it wrote.
        var root = Directory
            .CreateDirectory(Path.Combine(AHost("intact-key"), RecoveryContext.DirectoryName))
            .FullName;

        Assert.NotNull(RecoveryCapability.For(root));
        Assert.NotNull(RecoveryCapability.For(root));
    }

    [SkippableFact]
    public void TheRecoveryRoot_ShouldEndUpGrantingNobodyButThisAccount()
    {
        // R19-REC01, the Windows half. Checking the inherited ACL was never going to work: a
        // directory made under the user's temp root inherits whatever that root grants, and on an
        // ordinary machine that is several identities that are not the owner. So the ACL is set,
        // and this is the assertion that it really was.
        Skip.IfNot(OperatingSystem.IsWindows(), "Windows ACLs.");

        var host = AHost("acl-root");
        Assert.NotNull(RecoveryContext.For(host).Capability);

        if (OperatingSystem.IsWindows()) AssertPrivate(Path.Combine(host, RecoveryContext.DirectoryName));
    }

    [SkippableFact]
    public void ARecoveryRootWidenedAfterwards_ShouldBeNarrowedAgain()
    {
        // Widening it is something an attacker with write access to the directory can do at any
        // time, so the answer is to put it back on every start rather than to notice once.
        Skip.IfNot(OperatingSystem.IsWindows(), "Windows ACLs.");

        var host = AHost("widened-root");
        Assert.NotNull(RecoveryContext.For(host).Capability);

        var root = Path.Combine(host, RecoveryContext.DirectoryName);
        if (OperatingSystem.IsWindows()) Widen(root);

        Assert.NotNull(RecoveryCapability.For(root));
        if (OperatingSystem.IsWindows()) AssertPrivate(root);
    }

    /// <summary>Grants Everyone full control of a directory.</summary>
    /// <param name="directory">The directory to widen.</param>
    [SupportedOSPlatform("windows")]
    private static void Widen(string directory)
    {
        var info = new DirectoryInfo(directory);
        var security = info.GetAccessControl();
        security.AddAccessRule(new FileSystemAccessRule(
            new SecurityIdentifier(WellKnownSidType.WorldSid, null),
            FileSystemRights.FullControl, InheritanceFlags.ContainerInherit
                                          | InheritanceFlags.ObjectInherit,
            PropagationFlags.None, AccessControlType.Allow));
        info.SetAccessControl(security);
    }

    /// <summary>Asserts a directory inherits nothing and grants only the three trusted identities.</summary>
    /// <param name="directory">The directory to inspect.</param>
    [SupportedOSPlatform("windows")]
    private static void AssertPrivate(string directory)
    {
        var security = new DirectoryInfo(directory).GetAccessControl();

        Assert.True(security.AreAccessRulesProtected,
            "the recovery root still inherits its parent's grants");

        using var identity = WindowsIdentity.GetCurrent();
        var allowed = new[]
        {
            identity.User!,
            new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
            new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null)
        };

        foreach (FileSystemAccessRule rule in
                 security.GetAccessRules(true, true, typeof(SecurityIdentifier)))
        {
            if (rule.AccessControlType != AccessControlType.Allow) continue;

            Assert.Contains(allowed, sid => sid.Equals(rule.IdentityReference));
        }
    }

    [SkippableFact]
    public void AnOrdinaryUnixAccount_ShouldStillEstablishAKey()
    {
        // R20-REC04 refuses the mode inference for root. This is the control on the platform
        // where the check runs: an ordinary account is not refused. Running *as* root is not
        // exercised here — the CI account is not root — and is listed as such rather than
        // claimed.
        Skip.If(OperatingSystem.IsWindows(), "The uid check is Unix-only.");

        var host = AHost("unix-ordinary");
        Assert.NotNull(RecoveryContext.For(host).Capability);
    }

    [SkippableFact]
    public void AHostWhoseKeyFailedOnce_ShouldBeAbleToEstablishItLater()
    {
        // R20-REC07. `Established.GetOrAdd` cached whatever the first attempt produced, including
        // a context with no capability. A root that was briefly unusable — here, a link that is
        // later replaced by a real directory — stayed "no key" for the life of the process.
        var host = AHost("transient-host");
        var elsewhere = Directory.CreateDirectory(Path.Combine(TestDir, "elsewhere-t")).FullName;
        var root = Path.Combine(host, RecoveryContext.DirectoryName);

        Skip.IfNot(MidChainLinkFixture.TryCreateDirectoryLink(root, elsewhere),
            "This machine cannot create a directory link.");
        Assert.Null(RecoveryContext.For(host).Capability);

        // The condition clears: the link is removed and nothing is at the name.
        Directory.Delete(root, false);
        Assert.False(Directory.Exists(root));

        Assert.NotNull(RecoveryContext.For(host).Capability);
    }
}
