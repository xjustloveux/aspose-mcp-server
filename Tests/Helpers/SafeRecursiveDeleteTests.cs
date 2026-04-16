using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Tests for <see cref="SecurityHelper.SafeRecursiveDelete" /> covering symlink-safe
///     recursive deletion behaviour.
///     Platform gating: same as <see cref="ResolveAndEnsureWithinAllowlistTests" /> —
///     symlink-dependent cases are guarded by <c>[SkippableFact]</c>.
/// </summary>
public class SafeRecursiveDeleteTests : IDisposable
{
    /// <summary>Whether the current environment supports symlink creation.</summary>
    private static readonly bool SymlinksAvailable;

    private readonly TempScope _inside;
    private readonly TempScope _outside;

    static SafeRecursiveDeleteTests()
    {
        using var probe = SymlinkFixture.AllowlistedTempRoot();
        var probeLink = Path.Combine(probe.Root, "probe_link");
        var probeTarget = Path.Combine(probe.Root, "probe_target.txt");
        File.WriteAllText(probeTarget, "probe");
        SymlinksAvailable = SymlinkFixture.TryCreateFileSymlink(probeLink, probeTarget);
    }

    /// <summary>Creates two isolated temp scopes for inside/outside the allowlist.</summary>
    public SafeRecursiveDeleteTests()
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
    public void SafeRecursiveDelete_NormalTree_RemovesAllContentsAndRoot()
    {
        var root = InsidePath("normal_tree");
        var sub = Path.Combine(root, "subdir");
        var file1 = Path.Combine(root, "file1.txt");
        var file2 = Path.Combine(sub, "file2.txt");
        Directory.CreateDirectory(sub);
        File.WriteAllText(file1, "a");
        File.WriteAllText(file2, "b");

        SecurityHelper.SafeRecursiveDelete(root, Allowlist, nameof(root));

        Assert.False(Directory.Exists(root));
        Assert.False(File.Exists(file1));
        Assert.False(File.Exists(file2));
    }

    [SkippableFact]
    public void SafeRecursiveDelete_FileSymlinkOutside_RemovesLinkNotTarget()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var target = OutsidePath("preserved_target.txt");
        File.WriteAllText(target, "precious");

        var root = InsidePath("tree_with_file_link");
        Directory.CreateDirectory(root);
        var link = Path.Combine(root, "evil_file_link.txt");
        File.CreateSymbolicLink(link, target);

        SecurityHelper.SafeRecursiveDelete(root, Allowlist, nameof(root));

        Assert.False(File.Exists(link));
        Assert.False(Directory.Exists(root));
        Assert.True(File.Exists(target));
        Assert.Equal("precious", File.ReadAllText(target));
    }

    [SkippableFact]
    public void SafeRecursiveDelete_DirSymlinkOutside_RemovesLinkDoesNotDescendTarget()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var targetDir = OutsidePath("precious_outside_dir");
        var targetFile = Path.Combine(targetDir, "do_not_delete.txt");
        Directory.CreateDirectory(targetDir);
        File.WriteAllText(targetFile, "precious");

        var root = InsidePath("tree_with_dir_link");
        Directory.CreateDirectory(root);
        var linkDir = Path.Combine(root, "evil_dir_link");
        Directory.CreateSymbolicLink(linkDir, targetDir);

        SecurityHelper.SafeRecursiveDelete(root, Allowlist, nameof(root));

        Assert.False(Directory.Exists(linkDir));
        Assert.False(Directory.Exists(root));
        Assert.True(Directory.Exists(targetDir));
        Assert.True(File.Exists(targetFile));
        Assert.Equal("precious", File.ReadAllText(targetFile));
    }

    [SkippableFact]
    public void SafeRecursiveDelete_RootIsSymlink_RemovesLinkNotTarget()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var targetDir = InsidePath("real_root_target");
        var fileInTarget = Path.Combine(targetDir, "keep_me.txt");
        Directory.CreateDirectory(targetDir);
        File.WriteAllText(fileInTarget, "keep");

        var linkRoot = InsidePath("root_link");
        Directory.CreateSymbolicLink(linkRoot, targetDir);

        SecurityHelper.SafeRecursiveDelete(linkRoot, Allowlist, nameof(linkRoot));

        Assert.False(Directory.Exists(linkRoot));
        Assert.True(Directory.Exists(targetDir));
        Assert.True(File.Exists(fileInTarget));
    }

    [Fact]
    public void SafeRecursiveDelete_DeepNestedTree_SucceedsWithinBound()
    {
        const int depth = 10;
        var root = InsidePath("deep_tree");
        var leaf = root;
        for (var i = 0; i < depth; i++)
            leaf = Path.Combine(leaf, $"level_{i}");
        Directory.CreateDirectory(leaf);
        File.WriteAllText(Path.Combine(leaf, "deep_file.txt"), "deep");

        var ex = Record.Exception(() =>
            SecurityHelper.SafeRecursiveDelete(root, Allowlist, nameof(root)));

        Assert.Null(ex);
        Assert.False(Directory.Exists(root));
    }

    [SkippableFact]
    public void SafeRecursiveDelete_SubdirIsSymlink_ReCheckAtRecursionEntryRemovesLinkOnly()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var protectedDir = InsidePath("protected_real_dir");
        var protectedFile = Path.Combine(protectedDir, "protected.txt");
        Directory.CreateDirectory(protectedDir);
        File.WriteAllText(protectedFile, "protected");

        var root = InsidePath("nv2_test_root");
        Directory.CreateDirectory(root);
        var symlinkSubdir = Path.Combine(root, "planted_link");
        Directory.CreateSymbolicLink(symlinkSubdir, protectedDir);

        SecurityHelper.SafeRecursiveDelete(root, Allowlist, nameof(root));

        Assert.False(Directory.Exists(root));
        Assert.True(Directory.Exists(protectedDir));
        Assert.True(File.Exists(protectedFile));
        Assert.Equal("protected", File.ReadAllText(protectedFile));
    }

    [Fact]
    public void SafeRecursiveDelete_PathOutsideAllowlist_ThrowsBeforeAnyIO()
    {
        var outsideRoot = OutsidePath("should_not_be_deleted");
        var sentinel = Path.Combine(outsideRoot, "sentinel.txt");
        Directory.CreateDirectory(outsideRoot);
        File.WriteAllText(sentinel, "sentinel");

        Assert.Throws<ArgumentException>(() =>
            SecurityHelper.SafeRecursiveDelete(outsideRoot, Allowlist, nameof(outsideRoot)));

        Assert.True(File.Exists(sentinel));
    }
}
