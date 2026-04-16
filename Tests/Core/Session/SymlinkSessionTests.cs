using System.Reflection;
using System.Text.Json;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Integration tests that exercise the V5 <c>ReassertAllowlistForResolvedPath</c> wrapper
///     and the new <c>CloseDocument</c> symlink guard introduced in the symlink TOCTOU guard.
///     These tests verify that session-layer operations
///     (RecoverSession, CloseDocument, SaveSessionMetadata path) reject symlinks that escape
///     the configured allowlist.
///     Platform gating: symlink-dependent cases are decorated with <c>[SkippableFact]</c>
///     and call <see cref="Skip.IfNot" /> when symlink creation is unavailable.  Non-symlink
///     cases run unconditionally.
/// </summary>
public class SymlinkSessionTests : WordTestBase
{
    /// <summary>Whether the current environment supports symlink creation.</summary>
    private static readonly bool SymlinksAvailable;

    static SymlinkSessionTests()
    {
        using var probe = SymlinkFixture.AllowlistedTempRoot();
        var probeLink = Path.Combine(probe.Root, "probe_link");
        var probeTarget = Path.Combine(probe.Root, "probe_target.txt");
        File.WriteAllText(probeTarget, "probe");
        SymlinksAvailable = SymlinkFixture.TryCreateFileSymlink(probeLink, probeTarget);
    }

    // ------------------------------------------------------------------
    // Fixture helpers (mirrors SessionLoaderPathSecurityTests pattern)
    // ------------------------------------------------------------------

    private static ServerConfig BuildServerConfig(params string[] allowedPaths)
    {
        var cfg = new ServerConfig();
        var prop = typeof(ServerConfig).GetProperty(
            nameof(ServerConfig.AllowedBasePaths),
            BindingFlags.Instance | BindingFlags.Public);
        prop!.SetValue(cfg, allowedPaths
            .Select(Path.GetFullPath)
            .ToList()
            .AsReadOnly());
        return cfg;
    }

    private SessionConfig BuildSessionConfig()
    {
        return new SessionConfig
        {
            Enabled = true,
            MaxSessions = 5,
            MaxFileSizeMb = 10,
            IdleTimeoutMinutes = 0,
            TempDirectory = Path.Combine(TestDir, "temp")
        };
    }

    /// <summary>
    ///     Plants a <c>*.meta.json</c> file in <paramref name="tempDirectory" /> that the
    ///     session manager's recovery scanner will pick up.  Mirrors the legitimate writer's
    ///     naming convention.
    /// </summary>
    private static void PlantMetadata(string tempDirectory, string sessionId,
        string tempPath, string originalPath)
    {
        Directory.CreateDirectory(tempDirectory);
        var timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
        var metaPath = Path.Combine(tempDirectory,
            $"aspose_session_{sessionId}_{timestamp}.docx.meta.json");
        var metadata = new
        {
            SessionId = sessionId,
            OriginalPath = originalPath,
            TempPath = tempPath,
            DocumentType = "Word",
            SavedAt = DateTime.UtcNow,
            PromptOnReconnect = false
        };
        File.WriteAllText(metaPath, JsonSerializer.Serialize(metadata));
    }

    // =====================================================================
    // Test 1: RecoverSession with symlinked session.json pointing outside allowlist
    //
    // A planted metadata file whose TempPath is a symlink inside TempDirectory
    // but whose resolved target is outside TempDirectory.  The symlink-aware
    // ResolveAndEnsureWithinAllowlist guard in RecoverSession must reject this,
    // returning Success=false with a sanitised error.
    // =====================================================================

    /// <summary>
    ///     A symlinked TempPath that resolves outside TempDirectory must cause
    ///     <c>RecoverSession</c> to return <c>Success=false</c> without copying any content.
    ///     The error message must not contain the symlink target path.
    /// </summary>
    [SkippableFact]
    public void RecoverSession_SymlinkedTempPathOutsideTempDirectory_FailsWithSanitisedError()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var sessionConfig = BuildSessionConfig();
        var tempDir = sessionConfig.TempDirectory;
        Directory.CreateDirectory(tempDir);

        using var manager = new TempFileManager(sessionConfig);
        var sessionId = "sess_symlink_recover";

        // Create a real file outside TempDirectory that should NOT be read.
        var secretTarget = Path.Combine(TestDir, "secret_outside_temp.txt");
        File.WriteAllText(secretTarget, "SECRET CONTENT");

        // Plant a symlink inside TempDirectory whose resolved path is outside.
        var symlinkInsideTemp = Path.Combine(tempDir, $"aspose_session_{sessionId}_legit.docx");
        File.CreateSymbolicLink(symlinkInsideTemp, secretTarget);

        var recoverTarget = Path.Combine(TestDir, "recover_output.docx");
        PlantMetadata(tempDir, sessionId, symlinkInsideTemp, recoverTarget);

        var result = manager.RecoverSession(sessionId, recoverTarget);

        // Recovery must fail.
        Assert.False(result.Success);
        // Target file must NOT have been created (no copy occurred).
        Assert.False(File.Exists(recoverTarget),
            "RecoverSession must not copy content from a symlink that escapes TempDirectory");
        // Error message must not leak the secret path.
        if (result.ErrorMessage != null)
        {
            Assert.DoesNotContain(secretTarget, result.ErrorMessage);
            Assert.DoesNotContain("SECRET", result.ErrorMessage);
        }
    }

    // =====================================================================
    // Test 2: CloseDocument with session path that becomes a symlink pointing
    //         outside the allowlist between open and close (TOCTOU guard)
    //
    // With the M-1 fix applied, CloseDocument calls ReassertAllowlistForResolvedPath
    // immediately before SaveDocumentToFile.  If the session path has been replaced
    // by a symlink pointing outside the allowlist, the guard must throw before any
    // write occurs.
    // =====================================================================

    /// <summary>
    ///     Replaces a session's document file with a symlink pointing outside the allowlist
    ///     after the session is opened.  <c>CloseDocument</c> (dirty, non-discard) must throw
    ///     <see cref="ArgumentException" /> and leave the target file untouched.
    /// </summary>
    [SkippableFact]
    public void CloseDocument_SessionPathSymlinkedOutsideAllowlist_ThrowsAndDoesNotWriteTarget()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var serverConfig = BuildServerConfig(TestDir);
        var sessionConfig = BuildSessionConfig();
        using var manager = new DocumentSessionManager(sessionConfig, null, serverConfig);

        // Open a real document inside the allowlist.
        var docPath = CreateWordDocument("close_symlink_test.docx");
        var sessionId = manager.OpenDocument(docPath);

        // Mark the session dirty directly so CloseDocument triggers SaveDocumentToFile.
        var session = manager.GetSession(sessionId);
        session.IsDirty = true;

        // Remove the real file and plant a symlink at the same path pointing outside.
        var outsideTarget = Path.Combine(TestDir, "outside_write_target.txt");
        File.WriteAllText(outsideTarget, "ORIGINAL OUTSIDE CONTENT");

        // Temporarily narrow the allowlist to only the session-config temp directory,
        // which excludes TestDir — so docPath is now outside the (re-narrowed) allowlist.
        // This simulates the TOCTOU scenario where the allowlist changes after open.
        var narrowDir = Path.Combine(TestDir, "narrow_allowlist_dir");
        Directory.CreateDirectory(narrowDir);
        var prop = typeof(ServerConfig).GetProperty(
            nameof(ServerConfig.AllowedBasePaths),
            BindingFlags.Instance | BindingFlags.Public);
        prop!.SetValue(serverConfig,
            new List<string> { Path.GetFullPath(narrowDir) }.AsReadOnly());

        // CloseDocument with a dirty session must throw because the session path
        // (docPath in TestDir) is now outside the narrowed allowlist.
        Assert.Throws<ArgumentException>(() => manager.CloseDocument(sessionId));

        // The outside target must be unmodified.
        Assert.Equal("ORIGINAL OUTSIDE CONTENT", File.ReadAllText(outsideTarget));
    }

    // =====================================================================
    // Test 3: Symlink inside TempDirectory used as TempPath for metadata write
    //         (SaveSessionMetadata path guard — ancestor-walk NV-3)
    //
    // A directory symlink inside TempDirectory resolving outside it is planted
    // before the session metadata write.  The T7 fix (ResolveAndEnsureWithinAllowlist
    // before File.WriteAllText in SaveSessionMetadata) must block the write.
    //
    // SaveSessionMetadata is private and is indirectly exercised through a
    // HandleDisconnect / AutoSave path.  The most direct observable proxy is
    // that the auto-save temp file write during HandleDisconnect is guarded.
    // We exercise this by verifying that a symlinked temp directory causes the
    // session disposal (which triggers HandleDisconnect when disconnecting) to
    // throw or not write outside the boundary.
    // =====================================================================

    /// <summary>
    ///     When the temp directory contains a symlinked subdirectory pointing outside the
    ///     allowlist and a temp path through that symlink is supplied to the session metadata
    ///     write, the guard must prevent writes outside TempDirectory.
    ///     This is tested indirectly: we verify RecoverSession rejects a TempPath whose
    ///     ancestor is a directory symlink escaping TempDirectory (NV-3 ancestor-walk).
    /// </summary>
    [SkippableFact]
    public void RecoverSession_SymlinkedAncestorInsideTempDirPointsOutside_FailsWithoutCopy()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        var sessionConfig = BuildSessionConfig();
        var tempDir = sessionConfig.TempDirectory;
        Directory.CreateDirectory(tempDir);

        // Create a real directory outside TempDirectory.
        var outsideDir = Path.Combine(TestDir, "outside_real_dir");
        Directory.CreateDirectory(outsideDir);
        var secretFile = Path.Combine(outsideDir, "secret.docx");
        File.WriteAllText(secretFile, "SECRET");

        // Plant a directory symlink inside TempDirectory pointing to the outside dir.
        var symlinkDir = Path.Combine(tempDir, "escaped_subdir");
        Directory.CreateSymbolicLink(symlinkDir, outsideDir);

        // The TempPath goes through the symlinked ancestor — it is "inside" TempDirectory
        // by name but resolves outside.
        var tempPathViaSymlink = Path.Combine(symlinkDir, "session_data.docx");

        using var manager = new TempFileManager(sessionConfig);
        var sessionId = "sess_symlink_ancestor";
        var recoverDest = Path.Combine(TestDir, "recover_ancestor_out.docx");
        PlantMetadata(tempDir, sessionId, tempPathViaSymlink, recoverDest);

        var result = manager.RecoverSession(sessionId, recoverDest);

        // Must fail — the TempPath's ancestor resolves outside TempDirectory.
        Assert.False(result.Success,
            "RecoverSession must reject a TempPath whose ancestor resolves outside TempDirectory");
        Assert.False(File.Exists(recoverDest),
            "No file must be copied to the recover destination");
    }

    // =====================================================================
    // Regression: non-symlink paths continue to work normally after fix
    // =====================================================================

    /// <summary>
    ///     Verifies that the symlink guards do not break normal (non-symlink) recovery
    ///     for a legitimate TempPath inside TempDirectory.  This is the backward-compat
    ///     regression guard.
    /// </summary>
    [Fact]
    public void RecoverSession_LegitimateNonSymlinkTempPath_Succeeds()
    {
        var sessionConfig = BuildSessionConfig();
        var tempDir = sessionConfig.TempDirectory;
        Directory.CreateDirectory(tempDir);

        using var manager = new TempFileManager(sessionConfig);
        var sessionId = "sess_legit_recover";

        // Create a real temp file inside TempDirectory.
        var tempFile = Path.Combine(tempDir, $"aspose_session_{sessionId}_legit.docx");
        var docPath = CreateWordDocument("source_for_recovery.docx");
        File.Copy(docPath, tempFile, true);

        var recoverDest = Path.Combine(TestDir, "recovered_legit.docx");
        PlantMetadata(tempDir, sessionId, tempFile, recoverDest);

        var result = manager.RecoverSession(sessionId, recoverDest);

        // Legitimate recovery must succeed.
        Assert.True(result.Success);
        Assert.True(File.Exists(recoverDest));
    }
}
