using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R20-CNV01 and R20-CNV02: what the input copy refuses, and where its bytes come from.
///     <para>
///         The copy went under <c>recovery.Directory</c> whether or not the host had a capability.
///         A null capability means the root could not be shown to be private, and a copy of the
///         caller's whole input landed there anyway — the privacy failure became a data sink.
///     </para>
///     <para>
///         The bytes came from <c>File.Copy(source, copy)</c>, a pathname resolved again after the
///         entry check. They now come from one handle opened on the source resolved immediately
///         before; the window that remains is the same one <c>R18-SEC03</c> could not close, and
///         is documented on the method rather than claimed shut.
///     </para>
/// </summary>
public class ImmutableInputCopyTests : TestBase
{
    [SkippableFact]
    public void AHostWithNoCapability_ShouldNotStageAnInputAnywhere()
    {
        // R20-CNV01. A recovery root that is a link is one this server refuses to establish a key
        // in (R19-REC01), which is the cheapest way to get a context whose capability is null.
        var host = Directory.CreateDirectory(Path.Combine(TestDir, "capless-host")).FullName;
        var elsewhere = Directory.CreateDirectory(Path.Combine(TestDir, "elsewhere")).FullName;
        Skip.IfNot(MidChainLinkFixture.TryCreateDirectoryLink(
                Path.Combine(host, RecoveryContext.DirectoryName), elsewhere),
            "This machine cannot create a directory link.");

        var recovery = RecoveryContext.For(host);
        Assert.Null(recovery.Capability);

        var input = CreateTestFilePath("input.mht");
        File.WriteAllText(input, "MIME-Version: 1.0");

        Assert.Throws<InvalidOperationException>(() =>
            ImmutableInputCopy.Of(input, recovery, [TestDir]));

        // Nothing was written anywhere the refused host could reach.
        Assert.Empty(Directory.GetFiles(elsewhere, "*", SearchOption.AllDirectories));
        Assert.False(Directory.Exists(Path.Combine(recovery.Directory,
            ImmutableInputCopy.DirectoryName)));
    }

    [Fact]
    public void AnInputOutsideTheAllowlist_ShouldBeRefusedAtTheCopy()
    {
        // R20-CNV02. The entry resolved the path once; the copy resolves it again immediately
        // before opening it, so a check that was true a moment ago is re-asked of the file about
        // to be read.
        var outside = Directory.CreateDirectory(Path.Combine(TestDir, "outside")).FullName;
        var allowed = Directory.CreateDirectory(Path.Combine(TestDir, "allowed")).FullName;

        var input = Path.Combine(outside, "not_ours.mht");
        File.WriteAllText(input, "MIME-Version: 1.0");

        Assert.Throws<ArgumentException>(() =>
            ImmutableInputCopy.Of(input, Recovery, [allowed]));
    }

    [Fact]
    public void TheCopy_ShouldHoldTheResolvedSourcesBytesAndBeRemovedOnDispose()
    {
        // The control: an ordinary input is copied byte for byte from one open on the resolved
        // source, lives under this host's private root, and goes when the conversion does.
        var input = CreateTestFilePath("ours.mht");
        File.WriteAllText(input, "MIME-Version: 1.0\r\nContent-Type: text/html\r\n\r\n<p>hi</p>");

        string copyPath;
        using (var copy = ImmutableInputCopy.Of(input, Recovery, [TestDir]))
        {
            copyPath = copy.Path;
            Assert.StartsWith(Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName),
                copyPath, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(File.ReadAllBytes(input), File.ReadAllBytes(copyPath));
        }

        Assert.False(File.Exists(copyPath));
    }

    [Fact]
    public void ACopy_ShouldOccupyItsOwnTwoLevelNonceBucketUnderTheStagingDirectory()
    {
        // Physical shards let the startup sweeper rotate through bounded namespaces without
        // replaying an ever-growing flat prefix merely to reach its cursor.
        var input = CreateTestFilePath("sharded.mht");
        File.WriteAllText(input, "MIME-Version: 1.0");

        using var copy = ImmutableInputCopy.Of(input, Recovery, [TestDir]);

        var bucket = Directory.GetParent(copy.Path);
        Assert.NotNull(bucket);
        Assert.Matches("^[0-9a-f]{2}$", bucket.Name);
        var shard = bucket.Parent;
        Assert.NotNull(shard);
        Assert.Matches("^[0-9a-f]{2}$", shard.Name);
        Assert.Equal(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName),
            shard.Parent!.FullName,
            true);
    }
}
