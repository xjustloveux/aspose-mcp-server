using System.Reflection;
using System.Text.Json;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Regression tests for bug 20260415-session-loader-path. The fix wires
///     <see cref="AsposeMcpServer.Helpers.SecurityHelper.ValidateUserPath" /> into the three user-facing
///     trust boundaries of the session subsystem (OpenDocument / SaveDocument /
///     RecoverSession), closes the CRITICAL-1 metadata-driven File.Copy / File.Delete
///     primitives in <see cref="TempFileManager" />, re-runs the allowlist at
///     time-of-write to cover the HIGH-1 mutable-allowlist race, and strips
///     user-controlled paths out of exception messages (MEDIUM-1).
///     These xUnit cases are the behavioural oracle for those invariants; they also
///     act as regression guards against the peer "half-fix" pattern of
///     <c>20260415-excel-split-path-traversal</c>.
/// </summary>
public class SessionLoaderPathSecurityTests : WordTestBase
{
    // =====================================================================
    // Path-traversal payloads — OpenDocument / SaveDocument / RecoverSession
    //
    // Each trust boundary must run ValidateUserPath, which rejects:
    //   - `..` sequences (traversal),
    //   - NUL byte (invalid path character),
    //   - absolute paths outside the allowlist when one is configured.
    // =====================================================================

    public static TheoryData<string> TraversalPayloads =>
    [
        // Relative traversal payload — shape-rejected by ValidateFilePath.
        "../../etc/passwd",
        // NUL-byte injection — rejected as invalid path character.
        "file\0.docx"
    ];
    // ---------------------------------------------------------------------
    // Fixture helpers — build ServerConfigs with a given allowlist without
    // going through the command-line parser. AllowedBasePaths' setter is
    // private, so reflection is the least-invasive way to construct one in
    // tests (matches the approach other security-focused test suites use).
    // ---------------------------------------------------------------------

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
    ///     Plants a <c>*.meta.json</c> file in <paramref name="tempDirectory" /> pointing
    ///     <c>TempPath</c> and <c>OriginalPath</c> at <paramref name="tempPath" /> and
    ///     <paramref name="originalPath" /> respectively. Mirrors the legitimate writer's
    ///     naming convention so the manager's <c>Directory.GetFiles(...pattern)</c> call
    ///     finds it, simulating an attacker who can write to <c>Config.TempDirectory</c>.
    /// </summary>
    private static string PlantMetadata(string tempDirectory, string sessionId,
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
        return metaPath;
    }

    // =====================================================================
    // CRITICAL-1 — Planted-metadata attack: TempPath outside TempDirectory
    //
    // Pre-fix: ReadMetadata deserialized the planted JSON and File.Copy /
    // File.Delete operated on the attacker-chosen TempPath (read-anywhere,
    // delete-anywhere primitives reachable via the MCP `recover`,
    // `delete_temp`, and `cleanup` ops).
    //
    // Post-fix: an IsPathUnder(TempPath, Config.TempDirectory) prefix check
    // gates File.Copy and File.Delete. A TempPath outside TempDirectory must
    // cause recover to return Success=false with a generic error, and must
    // leave the foreign file on disk untouched for delete / cleanup paths.
    // =====================================================================

    [Fact]
    public void RecoverSession_MetadataTempPathOutsideTempDirectory_ShouldReject()
    {
        using var manager = new TempFileManager(BuildSessionConfig());
        var sessionId = "sess_planted1";
        // Point the planted TempPath at a real file outside TempDirectory; if the guard
        // is missing, File.Copy would succeed and leak this file to OriginalPath.
        var foreignFile = CreateWordDocument("foreign_secret.docx");
        var tempDir = Path.Combine(TestDir, "temp");
        var recoverTarget = CreateTestFilePath("recover_out.docx");
        PlantMetadata(tempDir, sessionId, foreignFile, recoverTarget);

        var result = manager.RecoverSession(sessionId, recoverTarget);

        Assert.False(result.Success);
        // Generic message — do not assume specific wording, but confirm the
        // rejection is path-location related and that recovery did not happen.
        Assert.False(File.Exists(recoverTarget), "Recovery must not copy foreign file out of TempDirectory");
    }

    [Fact]
    public void RecoverSession_MetadataOriginalPathOutsideAllowlist_ShouldReject()
    {
        // Allowlist restricts writes to TestDir; planted OriginalPath points elsewhere.
        // With no user-supplied targetPath, destination = metadata.OriginalPath, which
        // must be re-validated through ValidateUserPath (CRITICAL-1 branch).
        var serverCfg = BuildServerConfig(TestDir);
        var sessionCfg = BuildSessionConfig();
        using var manager = new TempFileManager(sessionCfg, null, serverCfg);
        var sessionId = "sess_planted2";

        // Legitimate TempPath under TempDirectory, but OriginalPath points at /tmp —
        // outside the allowlist. Null targetPath forces the fallback branch.
        var tempDir = Path.Combine(TestDir, "temp");
        Directory.CreateDirectory(tempDir);
        var fakeTempPath = Path.Combine(tempDir, $"aspose_session_{sessionId}_xx.docx");
        File.WriteAllText(fakeTempPath, "payload");
        var outsideAllowlist = Path.Combine(Path.GetTempPath(), "pwned_original.docx");
        PlantMetadata(tempDir, sessionId, fakeTempPath, outsideAllowlist);

        var ex = Record.Exception(() => manager.RecoverSession(sessionId));

        // Either an ArgumentException propagates out (preferred) or the operation
        // returns Success=false — both are acceptable; what is NOT acceptable is a
        // successful copy to outsideAllowlist.
        Assert.False(File.Exists(outsideAllowlist),
            "Recovery must not write to a path outside the allowlist");
        if (ex != null)
            Assert.IsType<ArgumentException>(ex);
    }

    [Fact]
    public void DeleteTempSession_MetadataTempPathOutsideTempDirectory_ShouldNotDeleteForeignFile()
    {
        using var manager = new TempFileManager(BuildSessionConfig());
        var sessionId = "sess_planted3";
        // A real file outside TempDirectory that must survive delete_temp.
        var foreignFile = CreateWordDocument("foreign_survives.docx");
        var tempDir = Path.Combine(TestDir, "temp");
        PlantMetadata(tempDir, sessionId, foreignFile, foreignFile);

        // delete_temp — reaches DeleteTempFileSet, which must NOT delete foreignFile.
        manager.DeleteTempSession(sessionId);

        Assert.True(File.Exists(foreignFile),
            "DeleteTempFileSet must refuse to File.Delete paths outside TempDirectory");
    }

    [Fact]
    public void CleanupExpiredFiles_MetadataTempPathOutsideTempDirectory_ShouldNotDeleteForeignFile()
    {
        // Expired planted metadata -> CleanupExpiredFiles -> DeleteTempFileSet.
        var sessionCfg = BuildSessionConfig();
        sessionCfg.TempRetentionHours = 1;
        using var manager = new TempFileManager(sessionCfg);
        var sessionId = "sess_planted4";
        var foreignFile = CreateWordDocument("foreign_cleanup.docx");
        var tempDir = Path.Combine(TestDir, "temp");
        var metaPath = PlantMetadata(tempDir, sessionId, foreignFile, foreignFile);
        // Backdate metadata so the cleanup scanner treats it as expired.
        File.SetLastWriteTimeUtc(metaPath, DateTime.UtcNow.AddHours(-48));

        manager.CleanupExpiredFiles();

        Assert.True(File.Exists(foreignFile),
            "Cleanup must refuse to File.Delete paths outside TempDirectory");
    }

    [Theory]
    [MemberData(nameof(TraversalPayloads))]
    public void OpenDocument_TraversalOrNulBytePayload_ShouldThrowArgumentException(string payload)
    {
        using var manager = new DocumentSessionManager(BuildSessionConfig());

        Assert.Throws<ArgumentException>(() => manager.OpenDocument(payload));
    }

    [Fact]
    public void OpenDocument_AbsolutePathOutsideAllowlist_ShouldThrowArgumentException()
    {
        // With an allowlist scoped to TestDir, an absolute path outside it
        // must be rejected even though the character/shape check passes.
        var serverCfg = BuildServerConfig(TestDir);
        using var manager = new DocumentSessionManager(
            BuildSessionConfig(), null, serverCfg);
        var outsidePath = Path.Combine(Path.GetTempPath(), "outside_allowlist.docx");

        Assert.Throws<ArgumentException>(() => manager.OpenDocument(outsidePath));
    }

    [Theory]
    [MemberData(nameof(TraversalPayloads))]
    public void SaveDocument_TraversalOrNulBytePayload_ShouldThrowArgumentException(string payload)
    {
        // SaveDocument validates outputPath BEFORE looking up the session, so we
        // can test the shape guard without needing a live session.
        using var manager = new DocumentSessionManager(BuildSessionConfig());

        Assert.Throws<ArgumentException>(() =>
            manager.SaveDocument("sess_does_not_matter", payload));
    }

    [Fact]
    public void SaveDocument_AbsolutePathOutsideAllowlist_ShouldThrowArgumentException()
    {
        var serverCfg = BuildServerConfig(TestDir);
        using var manager = new DocumentSessionManager(
            BuildSessionConfig(), null, serverCfg);
        var outsidePath = Path.Combine(Path.GetTempPath(), "outside_save.docx");

        Assert.Throws<ArgumentException>(() =>
            manager.SaveDocument("sess_does_not_matter", outsidePath));
    }

    [Theory]
    [MemberData(nameof(TraversalPayloads))]
    public void RecoverSession_TraversalOrNulByteTargetPath_ShouldThrowArgumentException(string payload)
    {
        // RecoverSession validates targetPath BEFORE touching disk; the guard
        // fires regardless of whether any metadata exists for the session ID.
        using var manager = new TempFileManager(BuildSessionConfig());

        Assert.Throws<ArgumentException>(() =>
            manager.RecoverSession("sess_does_not_matter", payload));
    }

    [Fact]
    public void RecoverSession_AbsoluteTargetPathOutsideAllowlist_ShouldThrowArgumentException()
    {
        var serverCfg = BuildServerConfig(TestDir);
        using var manager = new TempFileManager(BuildSessionConfig(), null, serverCfg);
        var outsidePath = Path.Combine(Path.GetTempPath(), "outside_recover.docx");

        Assert.Throws<ArgumentException>(() =>
            manager.RecoverSession("sess_does_not_matter", outsidePath));
    }

    // =====================================================================
    // HIGH-1 — Mutable-allowlist race: a session is opened under allowlist A,
    // then the admin narrows the allowlist (runtime config reload); a save
    // that falls back to session.Path must be re-checked against the CURRENT
    // allowlist, not the one in effect at open time.
    // =====================================================================

    [Fact]
    public void SaveDocument_AllowlistNarrowedAfterOpen_ShouldRejectFallbackSave()
    {
        // Open under a wide allowlist that includes TestDir.
        var serverCfg = BuildServerConfig(TestDir);
        using var manager = new DocumentSessionManager(
            BuildSessionConfig(), null, serverCfg);
        var docPath = CreateWordDocument("allowlist_race.docx");
        var sessionId = manager.OpenDocument(docPath);

        // Admin narrows the allowlist to a sibling directory that does NOT
        // contain docPath. session.Path is unchanged — it was captured at
        // open time — so the time-of-write re-check must fire on resolved
        // savePath and throw.
        var otherDir = Path.Combine(TestDir, "other");
        Directory.CreateDirectory(otherDir);
        var prop = typeof(ServerConfig).GetProperty(
            nameof(ServerConfig.AllowedBasePaths),
            BindingFlags.Instance | BindingFlags.Public);
        prop!.SetValue(serverCfg,
            new List<string> { Path.GetFullPath(otherDir) }.AsReadOnly());

        // Null outputPath -> save falls back to session.Path, which is now
        // outside the current allowlist.
        Assert.Throws<ArgumentException>(() => manager.SaveDocument(sessionId));
    }

    // =====================================================================
    // MEDIUM-1 — Error-message sanitization. The exception surfaced to the
    // caller must not echo attacker-controlled or internal paths (charter
    // §5). We assert on the error text rather than just the exception type.
    // =====================================================================

    [Fact]
    public void OpenDocument_FileNotFound_ShouldNotLeakPathInException()
    {
        // A valid-shape absolute path to a nonexistent file in TestDir — the
        // path survives ValidateUserPath and trips FileNotFoundException. The
        // post-fix exception message must NOT contain the probed path.
        using var manager = new DocumentSessionManager(BuildSessionConfig());
        var ghost = Path.Combine(TestDir, "ghost_secret_name.docx");

        var ex = Assert.Throws<FileNotFoundException>(() => manager.OpenDocument(ghost));

        Assert.DoesNotContain("ghost_secret_name", ex.Message);
        Assert.DoesNotContain(ghost, ex.Message);
    }

    [Fact]
    public void RecoverSession_MissingTempFile_ShouldNotLeakPathInErrorMessage()
    {
        // Plant a metadata file whose TempPath is legal (under TempDirectory)
        // but points at a file we never created. The generic "Temp file not
        // found" branch must NOT interpolate the probed path into the result.
        using var manager = new TempFileManager(BuildSessionConfig());
        var sessionId = "sess_missingtemp";
        var tempDir = Path.Combine(TestDir, "temp");
        Directory.CreateDirectory(tempDir);
        var ghostTemp = Path.Combine(tempDir,
            $"aspose_session_{sessionId}_nonexistent_unique.docx");
        var originalPath = CreateTestFilePath("recover_missing_out.docx");
        PlantMetadata(tempDir, sessionId, ghostTemp, originalPath);

        var result = manager.RecoverSession(sessionId);

        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.DoesNotContain("nonexistent_unique", result.ErrorMessage);
        Assert.DoesNotContain(ghostTemp, result.ErrorMessage);
    }
}
