using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Regression guard for RB-02: <c>SecurityHelper.ResolveSymlinkChain</c> previously called
///     <see cref="FileSystemInfo.ResolveLinkTarget" /> only on the leaf (when it exists) or on the
///     nearest existing ancestor (when it does not). Neither call inspects the segments in between,
///     so a junction or directory symlink placed inside an allowed directory redirected both reads
///     and writes outside the allowlist while the lexical check still passed.
///     These tests place a directory link mid-chain and assert the resolved path is rejected.
/// </summary>
public class MidChainSymlinkTests : IDisposable
{
    private readonly string _allowedRoot;
    private readonly bool _linkCreated;
    private readonly string _linkPath;

    /// <summary>
    ///     Builds an allowed root that contains a directory link pointing at a directory outside it.
    /// </summary>
    public MidChainSymlinkTests()
    {
        var baseDir = Path.Combine(Path.GetTempPath(), "MidChain_" + Guid.NewGuid().ToString("N"));
        _allowedRoot = Path.Combine(baseDir, "allowed");
        var secretRoot = Path.Combine(baseDir, "secret");
        Directory.CreateDirectory(_allowedRoot);
        Directory.CreateDirectory(secretRoot);
        Directory.CreateDirectory(Path.Combine(secretRoot, "sub"));
        File.WriteAllText(Path.Combine(secretRoot, "secret.txt"), "classified");

        _linkPath = Path.Combine(_allowedRoot, "link");
        _linkCreated = MidChainLinkFixture.TryCreateDirectoryLink(_linkPath, secretRoot);
    }

    /// <summary>Allowlist containing only the allowed root.</summary>
    private IReadOnlyList<string> Allowed => [_allowedRoot];

    /// <inheritdoc />
    public void Dispose()
    {
        GC.SuppressFinalize(this);
        var baseDir = Path.GetDirectoryName(_allowedRoot);
        if (baseDir == null || !Directory.Exists(baseDir)) return;
        try
        {
            if (_linkCreated && Directory.Exists(_linkPath)) Directory.Delete(_linkPath);
            Directory.Delete(baseDir, true);
        }
        catch (IOException)
        {
            // Best-effort cleanup; a locked file must not fail the test run.
        }
    }

    [SkippableFact]
    public void ExistingFileBehindMidChainLink_ShouldBeRejected()
    {
        Skip.IfNot(_linkCreated, "Directory link creation is not permitted in this environment.");
        var path = Path.Combine(_linkPath, "secret.txt");
        Assert.True(File.Exists(path), "Fixture must expose the outside file through the link.");

        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(path, Allowed, "path"));
    }

    [SkippableFact]
    public void NonExistentWriteSinkBehindMidChainLink_ShouldBeRejected()
    {
        Skip.IfNot(_linkCreated, "Directory link creation is not permitted in this environment.");
        var path = Path.Combine(_linkPath, "sub", "new-file.txt");
        Assert.False(File.Exists(path));

        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(path, Allowed, "path"));
    }

    [SkippableFact]
    public void DeepNonExistentTailBehindMidChainLink_ShouldBeRejected()
    {
        Skip.IfNot(_linkCreated, "Directory link creation is not permitted in this environment.");
        var path = Path.Combine(_linkPath, "missing", "deeper", "file.txt");

        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(path, Allowed, "path"));
    }

    [Fact]
    public void PathWithoutLinks_ShouldStillResolveInsideAllowlist()
    {
        var inside = Path.Combine(_allowedRoot, "plain", "file.txt");
        Directory.CreateDirectory(Path.GetDirectoryName(inside)!);
        File.WriteAllText(inside, "ok");

        var resolved = SecurityHelper.ResolveAndEnsureWithinAllowlist(inside, Allowed, "path");

        Assert.Equal(Path.GetFullPath(inside), resolved);
    }

    [Fact]
    public void NonExistentPathWithoutLinks_ShouldPreserveTail()
    {
        var sink = Path.Combine(_allowedRoot, "does", "not", "exist.txt");

        var resolved = SecurityHelper.ResolveAndEnsureWithinAllowlist(sink, Allowed, "path");

        Assert.Equal(Path.GetFullPath(sink), resolved);
    }
}
