using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Tests for <see cref="SecurityHelper.ResolveAndEnsureWithinAllowlist" /> covering all
///     symlink and path-resolution scenarios introduced by the 20260415-symlink-toctou-sweep
///     fix.  Each test probes the helper in isolation (no Aspose dependency).
///     Platform gating: <see cref="File.CreateSymbolicLink" /> on Windows requires Developer
///     Mode or administrator privileges.  The fixture calls <see cref="SymlinkFixture.TryCreateFileSymlink" />
///     in a static constructor probe; tests that create symlinks are decorated with
///     <c>[SkippableFact]</c> / <c>[SkippableTheory]</c> and call <see cref="Skip.If" /> at
///     the top when symlink creation is not available.  Non-symlink cases run unconditionally.
/// </summary>
public class ResolveAndEnsureWithinAllowlistTests : IDisposable
{
    /// <summary>
    ///     Whether the current OS and privilege level support symlink creation.
    ///     Determined once per process by attempting a probe symlink in a temp directory.
    /// </summary>
    private static readonly bool SymlinksAvailable;

    private readonly TempScope _inside;
    private readonly TempScope _outside;

    static ResolveAndEnsureWithinAllowlistTests()
    {
        using var probe = SymlinkFixture.AllowlistedTempRoot();
        var probeLink = Path.Combine(probe.Root, "probe_link");
        var probeTarget = Path.Combine(probe.Root, "probe_target.txt");
        File.WriteAllText(probeTarget, "probe");
        SymlinksAvailable = SymlinkFixture.TryCreateFileSymlink(probeLink, probeTarget);
    }

    /// <summary>Initialises two temp scopes: one inside the allowlist, one outside.</summary>
    public ResolveAndEnsureWithinAllowlistTests()
    {
        _inside = SymlinkFixture.AllowlistedTempRoot();
        _outside = SymlinkFixture.AllowlistedTempRoot();
    }

    private IReadOnlyList<string> Allowlist => [_inside.Root];

    /// <inheritdoc />
    public void Dispose()
    {
        _inside.Dispose();
        _outside.Dispose();
    }

    private string InsidePath(string relative)
    {
        return Path.Combine(_inside.Root, relative);
    }

    private string OutsidePath(string relative)
    {
        return Path.Combine(_outside.Root, relative);
    }

    [Fact]
    public void ResolveAndEnsureWithinAllowlist_NonSymlinkInsideAllowlist_ReturnsNormalisedPath()
    {
        var filePath = InsidePath("normal.txt");
        File.WriteAllText(filePath, "content");

        var result = SecurityHelper.ResolveAndEnsureWithinAllowlist(filePath, Allowlist, "path");

        Assert.NotNull(result);
        Assert.Equal(Path.GetFullPath(filePath), result);
    }

    [Fact]
    public void ResolveAndEnsureWithinAllowlist_NonSymlinkOutsideAllowlist_ThrowsArgumentException()
    {
        var filePath = OutsidePath("forbidden.txt");
        File.WriteAllText(filePath, "content");

        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(filePath, Allowlist, "outputPath"));
    }

    [SkippableFact]
    public void ResolveAndEnsureWithinAllowlist_FilesSymlinkTargetInsideAllowlist_ReturnsResolvedPath()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var target = InsidePath("target_file.txt");
        File.WriteAllText(target, "data");
        var link = InsidePath("link_to_target.txt");
        File.CreateSymbolicLink(link, target);

        var result = SecurityHelper.ResolveAndEnsureWithinAllowlist(link, Allowlist, "path");

        Assert.Equal(Path.GetFullPath(target), result);
    }

    [SkippableFact]
    public void ResolveAndEnsureWithinAllowlist_DirSymlinkTargetInsideAllowlist_ReturnsResolvedPath()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var realDir = InsidePath("real_subdir");
        Directory.CreateDirectory(realDir);
        var fileUnder = Path.Combine(realDir, "file.txt");
        File.WriteAllText(fileUnder, "data");

        var linkDir = InsidePath("linked_subdir");
        Directory.CreateSymbolicLink(linkDir, realDir);
        var pathViaLink = Path.Combine(linkDir, "file.txt");

        var result = SecurityHelper.ResolveAndEnsureWithinAllowlist(pathViaLink, Allowlist, "path");

        Assert.StartsWith(Path.GetFullPath(_inside.Root), result);
    }

    [SkippableFact]
    public void ResolveAndEnsureWithinAllowlist_FilesSymlinkTargetOutsideAllowlist_ThrowsAndDoesNotLeakPath()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var target = OutsidePath("secret.txt");
        File.WriteAllText(target, "secret content");
        var link = InsidePath("evil_link.txt");
        File.CreateSymbolicLink(link, target);

        var ex = Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(link, Allowlist, "outputPath"));

        Assert.DoesNotContain(target, ex.Message);
        Assert.DoesNotContain(_outside.Root, ex.Message);
    }

    [SkippableFact]
    public void ResolveAndEnsureWithinAllowlist_SymlinkChainAllInsideAllowlist_FollowsToFinalTarget()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var c = InsidePath("chain_c.txt");
        File.WriteAllText(c, "final");
        var b = InsidePath("chain_b.txt");
        File.CreateSymbolicLink(b, c);
        var a = InsidePath("chain_a.txt");
        File.CreateSymbolicLink(a, b);

        var result = SecurityHelper.ResolveAndEnsureWithinAllowlist(a, Allowlist, "path");

        Assert.Equal(Path.GetFullPath(c), result);
    }

    [SkippableFact]
    public void ResolveAndEnsureWithinAllowlist_DanglingSymlinkTargetOutsideAllowlist_Throws()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var nonExistentOutside = OutsidePath("does_not_exist.txt");
        var link = InsidePath("dangling_evil.txt");
        File.CreateSymbolicLink(link, nonExistentOutside);

        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(link, Allowlist, "path"));
    }

    [SkippableFact]
    public void ResolveAndEnsureWithinAllowlist_NonExistentPathWithSymlinkedAncestorOutside_Throws()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var outsideDir = OutsidePath("outside_dir");
        Directory.CreateDirectory(outsideDir);
        var symlinkDir = InsidePath("symlinked_ancestor");
        Directory.CreateSymbolicLink(symlinkDir, outsideDir);

        var nonExistentOutput = Path.Combine(symlinkDir, "output.docx");

        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(nonExistentOutput, Allowlist, "outputPath"));
    }

    [Fact]
    public void ResolveAndEnsureWithinAllowlist_NonExistentPathWithLegitimateAncestor_ReturnsReconstructedPath()
    {
        var subDir = InsidePath("legit_subdir");
        Directory.CreateDirectory(subDir);
        var nonExistentOutput = Path.Combine(subDir, "output.docx");

        var result = SecurityHelper.ResolveAndEnsureWithinAllowlist(nonExistentOutput, Allowlist, "outputPath");

        Assert.NotNull(result);
        Assert.StartsWith(Path.GetFullPath(_inside.Root), result);
        Assert.EndsWith("output.docx", result);
    }

    [Fact]
    public void ResolveAndEnsureWithinAllowlist_FailedCheck_ParamNameInMessageNoPathLeak()
    {
        var filePath = OutsidePath("leaked_path_secret.txt");
        File.WriteAllText(filePath, "secret");

        var ex = Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(filePath, Allowlist, "mySpecialParam"));

        Assert.Contains("mySpecialParam", ex.Message);
        Assert.DoesNotContain(filePath, ex.Message);
        Assert.DoesNotContain(Path.GetFullPath(filePath), ex.Message);
        Assert.DoesNotContain(_outside.Root, ex.Message);
    }

    [SkippableFact]
    public void ResolveAndEnsureWithinAllowlist_CircularSymlink_ThrowsArgumentExceptionSanitised()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var a = InsidePath("circular_a.txt");
        var b = InsidePath("circular_b.txt");

        File.CreateSymbolicLink(b, a);
        File.CreateSymbolicLink(a, b);

        var ex = Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(a, Allowlist, "circularPath"));

        Assert.DoesNotContain(_inside.Root, ex.Message);
        Assert.Contains("circularPath", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ResolveAndEnsureWithinAllowlist_NullOrEmptyPath_ThrowsArgumentException(string? path)
    {
        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.ResolveAndEnsureWithinAllowlist(path!, Allowlist, "p"));
    }
}
