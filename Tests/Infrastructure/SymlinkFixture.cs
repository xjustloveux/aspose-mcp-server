namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Helpers for creating and managing symbolic links in tests.
///     All paths are scoped to <see cref="Path.GetTempPath" />-derived directories so no
///     real user paths are touched.  Each method is designed for use by test fixtures that
///     gate symlink cases with a probe-create check; callers should call
///     <see cref="TryCreateFileSymlink" /> in fixture setup and skip the test when it
///     returns <c>false</c> (Windows without Developer Mode / unprivileged CI).
/// </summary>
public static class SymlinkFixture
{
    /// <summary>
    ///     Attempts to create a file symbolic link at <paramref name="linkPath" /> pointing to
    ///     <paramref name="targetPath" />.  Returns <c>true</c> on success, <c>false</c> if the
    ///     OS or privilege level prevents symlink creation (e.g. Windows without Developer Mode).
    /// </summary>
    /// <param name="linkPath">Absolute path where the symlink entry should be created.</param>
    /// <param name="targetPath">Absolute or relative target of the symlink.</param>
    /// <returns><c>true</c> if the link was created; <c>false</c> otherwise.</returns>
    public static bool TryCreateFileSymlink(string linkPath, string targetPath)
    {
        try
        {
            File.CreateSymbolicLink(linkPath, targetPath);
            return File.Exists(linkPath) || new FileInfo(linkPath).LinkTarget != null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    ///     Creates a directory symbolic link at <paramref name="linkPath" /> pointing to
    ///     <paramref name="targetPath" />.  Throws if symlink creation fails (callers that
    ///     need graceful skip should gate with <see cref="TryCreateFileSymlink" /> first).
    /// </summary>
    /// <param name="linkPath">Absolute path where the directory symlink entry should be created.</param>
    /// <param name="targetPath">Absolute or relative target directory.</param>
    /// <exception cref="IOException">
    ///     Thrown when symlink creation is not permitted by the OS or privilege level.
    /// </exception>
    public static void CreateDirSymlink(string linkPath, string targetPath)
    {
        Directory.CreateSymbolicLink(linkPath, targetPath);
    }

    /// <summary>
    ///     Returns a disposable scope that creates a uniquely-named subdirectory under
    ///     <see cref="Path.GetTempPath" /> and deletes it (and all contents) on disposal.
    ///     The returned root is safe to use as an allowlist entry in tests.
    /// </summary>
    /// <returns>
    ///     A <see cref="TempScope" /> whose <see cref="TempScope.Root" /> is the created directory.
    /// </returns>
    public static TempScope AllowlistedTempRoot()
    {
        var root = Path.Combine(Path.GetTempPath(), "SymlinkTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return new TempScope(root);
    }
}
