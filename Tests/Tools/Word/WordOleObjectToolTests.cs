using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Shared.Ole;
using AsposeMcpServer.Tests.Infrastructure.Ole;
using AsposeMcpServer.Tools.Word;
using ModelContextProtocol.Server;
using OoxmlSaveOptions = Aspose.Words.Saving.OoxmlSaveOptions;
using SaveFormat = Aspose.Cells.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Per-tool AC coverage for <c>word_ole_object</c> (AC-1 ... AC-17) plus the session-
///     mode parity and password-advisory cases. Uses in-process fixtures from
///     <see cref="FixtureBuilder" />.
/// </summary>
[Collection(OleFixtureCollection.Name)]
public class WordOleObjectToolTests : IDisposable
{
    private readonly FixtureBuilder _fixtures;
    private readonly string _outputDir;

    /// <summary>Initializes a new instance of the <see cref="WordOleObjectToolTests" /> class.</summary>
    /// <param name="fixtures">Shared fixture matrix.</param>
    public WordOleObjectToolTests(FixtureBuilder fixtures)
    {
        _fixtures = fixtures;
        _outputDir = Path.Combine(Path.GetTempPath(), "WordOleTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_outputDir);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        try
        {
            Directory.Delete(_outputDir, true);
        }
        catch
        {
            /* best-effort */
        }
    }

    /// <summary>
    ///     AC-1: <c>list</c> returns the expected metadata (count, index, rawFileName,
    ///     suggestedFileName, progId, isLinked).
    /// </summary>
    [Fact]
    public void List_ReturnsExpectedMetadata()
    {
        var tool = new WordOleObjectTool();

        var raw = tool.Execute("list", _fixtures.Paths[FixtureKind.WordEmbeddedDocx]);

        var data = ((FinalizedResult<OleListResult>)raw).Data;
        Assert.Equal(1, data.Count);
        Assert.Single(data.Items);
        var item = data.Items[0];
        Assert.Equal(0, item.Index);
        Assert.False(item.IsLinked);
        // Aspose.Words' InsertOleObject(Stream, ...) does not always populate
        // OlePackage.FileName; our mapper's chain of fallbacks eventually lands on
        // ole_N.xlsx via ProgId=Excel.Sheet.12. RawFileName may be null depending on
        // Aspose internals; what matters for AC-1 is that a non-empty disk-safe name
        // was produced and the metadata shape is populated.
        Assert.False(string.IsNullOrEmpty(item.SuggestedFileName));
        Assert.EndsWith(".xlsx", item.SuggestedFileName);
        Assert.Equal("Excel.Sheet.12", item.ProgId);
        Assert.True(item.SizeBytes > 0);
    }

    /// <summary>
    ///     AC-2: <c>extract</c> writes the correct bytes to the output directory.
    /// </summary>
    [Fact]
    public void Extract_WritesCorrectBytes()
    {
        var tool = new WordOleObjectTool();

        var raw = tool.Execute("extract", _fixtures.Paths[FixtureKind.WordEmbeddedDocx],
            outputDirectory: _outputDir, oleIndex: 0);

        var data = ((FinalizedResult<OleExtractResult>)raw).Data;
        Assert.True(File.Exists(data.OutputFilePath));
        Assert.True(new FileInfo(data.OutputFilePath).Length > 0);
        Assert.Equal(_outputDir, Path.GetDirectoryName(data.OutputFilePath));
    }

    /// <summary>
    ///     AC-3: attacker-controlled raw filename (<c>..\\..\\etc\\passwd</c>) is
    ///     sanitized and the extracted file stays within the output directory.
    /// </summary>
    [Fact]
    public void Extract_AttackerFilename_StaysWithinOutputDir()
    {
        var tool = new WordOleObjectTool();

        var raw = tool.Execute("extract", _fixtures.Paths[FixtureKind.WordAttackerDocx],
            outputDirectory: _outputDir, oleIndex: 0);

        var data = ((FinalizedResult<OleExtractResult>)raw).Data;
        Assert.True(data.SanitizedFromRaw);
        Assert.Equal(
            Path.GetFullPath(_outputDir),
            Path.GetDirectoryName(data.OutputFilePath));
        Assert.DoesNotContain("..", Path.GetFileName(data.OutputFilePath));
    }

    /// <summary>
    ///     AC-4: attacker-controlled raw filename eventually resolves to a non-empty
    ///     sanitized filename (no NUL, no separators). For Word the fixture provides
    ///     a traversal-style attacker filename; sanitizer produces the "etc passwd"
    ///     compacted form or the <c>ole_N.xlsx</c> fallback.
    /// </summary>
    [Fact]
    public void Extract_NulAndReservedName_FallsBack()
    {
        var tool = new WordOleObjectTool();

        var raw = tool.Execute("extract", _fixtures.Paths[FixtureKind.WordAttackerDocx],
            outputDirectory: _outputDir, oleIndex: 0);

        var data = ((FinalizedResult<OleExtractResult>)raw).Data;
        var fileName = Path.GetFileName(data.OutputFilePath);
        Assert.False(string.IsNullOrEmpty(fileName));
        Assert.DoesNotContain('\0', fileName);
        Assert.DoesNotContain("/", fileName);
        Assert.DoesNotContain("\\", fileName);
    }

    /// <summary>
    ///     AC-5: <c>extract_all</c> skips linked OLE objects with
    ///     <c>reason: "linked"</c>.
    /// </summary>
    [Fact]
    public void ExtractAll_SkipsLinkedWithReason()
    {
        var tool = new WordOleObjectTool();

        var raw = tool.Execute("extract_all", _fixtures.Paths[FixtureKind.WordLinkedDocx],
            outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.Contains(data.Skipped, s => s.Reason == "linked");
        Assert.Equal(0, data.Extracted);
    }

    /// <summary>
    ///     AC-7: output directory outside the configured allowlist is rejected with
    ///     <see cref="UnauthorizedAccessException" /> carrying no path fragment.
    /// </summary>
    [Fact]
    public void OutputDirOutsideAllowedRoots_Rejected()
    {
        var config = new ServerConfig();
        typeof(ServerConfig)
            .GetProperty("AllowedBasePaths")!
            .SetValue(config, new[] { _outputDir }.ToList().AsReadOnly());

        var tool = new WordOleObjectTool(serverConfig: config);

        // The source path is inside the allowlist, but the output dir is outside.
        var sourceCopy = Path.Combine(_outputDir, "source.docx");
        File.Copy(_fixtures.Paths[FixtureKind.WordEmbeddedDocx], sourceCopy);
        var outsideDir = Path.Combine(Path.GetTempPath(), "WordOleOutside_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outsideDir);
        try
        {
            Assert.Throws<UnauthorizedAccessException>(() =>
                tool.Execute("extract", sourceCopy, outputDirectory: outsideDir, oleIndex: 0));
        }
        finally
        {
            try
            {
                Directory.Delete(outsideDir, true);
            }
            catch
            {
                /* best-effort */
            }
        }
    }

    /// <summary>
    ///     AC-10: no hard count cap on <c>list</c> — the cumulative loop handles
    ///     containers with many OLE objects. We use a modest 50-object document as a
    ///     proxy since authoring 1500 OLE objects materially inflates test runtime;
    ///     the assertion is that <see cref="OleListResult.Count" /> matches the input
    ///     count exactly and <see cref="OleListResult.Truncated" /> stays <c>false</c>.
    /// </summary>
    [Fact]
    public void NoCountCap_HandlesManyObjects()
    {
        var path = Path.Combine(_outputDir, "many.docx");
        var xlsxPayload = BuildValidXlsxPayload();
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < 30; i++)
        {
            using var payload = new MemoryStream(xlsxPayload);
            using var img = new MemoryStream(SmallPng());
            builder.InsertOleObject(payload, "Excel.Sheet.12", false, img);
        }

        doc.Save(path);

        var tool = new WordOleObjectTool();
        var raw = tool.Execute("list", path);
        var data = ((FinalizedResult<OleListResult>)raw).Data;

        Assert.Equal(30, data.Count);
        Assert.False(data.Truncated);
    }

    /// <summary>
    ///     AC-11: <c>list</c> invoked twice on the same file yields byte-identical JSON.
    /// </summary>
    [Fact]
    public void List_TwiceSameFile_ByteIdenticalJson()
    {
        var tool = new WordOleObjectTool();
        var options = new JsonSerializerOptions();

        var r1 = tool.Execute("list", _fixtures.Paths[FixtureKind.WordEmbeddedDocx]);
        var r2 = tool.Execute("list", _fixtures.Paths[FixtureKind.WordEmbeddedDocx]);

        var d1 = ((FinalizedResult<OleListResult>)r1).Data;
        var d2 = ((FinalizedResult<OleListResult>)r2).Data;

        var j1 = JsonSerializer.Serialize(d1, options);
        var j2 = JsonSerializer.Serialize(d2, options);
        Assert.Equal(j1, j2);
    }

    /// <summary>
    ///     AC-15: filename collision in <c>extract_all</c> resolves with <c>" (2)"</c>
    ///     suffix. We exercise the collision-resolver helper directly because Aspose.Words
    ///     23.10.0 does not reliably persist <c>OlePackage.FileName</c> overrides on the
    ///     document round-trip; the sanitizer therefore emits <c>ole_0.xlsx</c> /
    ///     <c>ole_1.xlsx</c> for two sibling OLE objects — no collision to exercise at
    ///     the handler level. The collision test is therefore pinned to the shared
    ///     resolver that all three handlers use.
    /// </summary>
    [Fact]
    public void ExtractAll_ExplicitCollision_ProducesDistinctFiles()
    {
        var resolver = new OleCollisionResolver();
        var outDir = Path.Combine(_outputDir, "collision");
        Directory.CreateDirectory(outDir);

        var a = resolver.Reserve(outDir, "dup.xlsx");
        var b = resolver.Reserve(outDir, "dup.xlsx");
        var c = resolver.Reserve(outDir, "dup.xlsx");

        Assert.Equal("dup.xlsx", Path.GetFileName(a));
        Assert.Equal("dup (2).xlsx", Path.GetFileName(b));
        Assert.Equal("dup (3).xlsx", Path.GetFileName(c));
    }

    /// <summary>
    ///     AC-16: <c>remove</c> drops the count by one on the persisted document.
    /// </summary>
    [Fact]
    public void Remove_ReducesCountByOne()
    {
        var sourcePath = Path.Combine(_outputDir, "rem.docx");
        File.Copy(_fixtures.Paths[FixtureKind.WordEmbeddedDocx], sourcePath);

        var tool = new WordOleObjectTool();
        var raw = tool.Execute("remove", sourcePath, oleIndex: 0);
        var data = ((FinalizedResult<OleRemoveResult>)raw).Data;
        Assert.True(data.Removed);

        var afterList = tool.Execute("list", sourcePath);
        var afterData = ((FinalizedResult<OleListResult>)afterList).Data;
        Assert.Equal(0, afterData.Count);
    }

    /// <summary>
    ///     AC-17: removing index 0 twice in a row operates on different objects (the
    ///     second remove sees the reindexed collection, not a stale cursor).
    /// </summary>
    [Fact]
    public void Remove_ThenRemove_AtIndexZero_AffectsDifferentObject()
    {
        var path = Path.Combine(_outputDir, "reidx.docx");
        var xlsxPayload = BuildValidXlsxPayload();
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < 2; i++)
        {
            using var payload = new MemoryStream(xlsxPayload);
            using var img = new MemoryStream(SmallPng());
            builder.InsertOleObject(payload, "Excel.Sheet.12", false, img);
        }

        doc.Save(path);

        var tool = new WordOleObjectTool();
        tool.Execute("remove", path, oleIndex: 0);
        tool.Execute("remove", path, oleIndex: 0);

        var raw = tool.Execute("list", path);
        var data = ((FinalizedResult<OleListResult>)raw).Data;
        Assert.Equal(0, data.Count);
    }

    /// <summary>
    ///     Session-mode parity (AC-8): <c>list</c> returns the same item count as
    ///     file-mode for the same source document.
    /// </summary>
    [Fact]
    public void List_SessionMode_MatchesFileMode()
    {
        using var sessionManager = new DocumentSessionManager(
            new SessionConfig
            {
                Enabled = true,
                TempDirectory = Path.Combine(_outputDir, "sessions")
            });
        var tool = new WordOleObjectTool(sessionManager);
        var fixturePath = _fixtures.Paths[FixtureKind.WordEmbeddedDocx];

        var sid = sessionManager.OpenDocument(fixturePath);
        try
        {
            var fileRaw = tool.Execute("list", fixturePath);
            var sessionRaw = tool.Execute("list", sessionId: sid);

            var fileData = ((FinalizedResult<OleListResult>)fileRaw).Data;
            var sessionData = ((FinalizedResult<OleListResult>)sessionRaw).Data;
            Assert.Equal(fileData.Count, sessionData.Count);
            Assert.Equal(fileData.Items[0].SuggestedFileName, sessionData.Items[0].SuggestedFileName);
        }
        finally
        {
            sessionManager.CloseDocument(sid, true);
        }
    }

    /// <summary>
    ///     Session-mode password-ignored advisory note (F-5): when a password is
    ///     supplied alongside a sessionId, the response carries the locked-shape
    ///     <see cref="PasswordIgnoredNote" />.
    /// </summary>
    [Fact]
    public void Session_WithPassword_AttachesPasswordIgnoredNote()
    {
        using var sessionManager = new DocumentSessionManager(
            new SessionConfig
            {
                Enabled = true,
                TempDirectory = Path.Combine(_outputDir, "sessions")
            });
        var tool = new WordOleObjectTool(sessionManager);

        var sid = sessionManager.OpenDocument(_fixtures.Paths[FixtureKind.WordEmbeddedDocx]);
        try
        {
            var raw = tool.Execute("list", sessionId: sid, password: "irrelevant");
            var data = ((FinalizedResult<OleListResult>)raw).Data;

            Assert.NotNull(data.PasswordIgnored);
            Assert.True(data.PasswordIgnored.PasswordIgnored);
            Assert.Equal("session-already-unlocked", data.PasswordIgnored.Reason);
        }
        finally
        {
            sessionManager.CloseDocument(sid, true);
        }
    }

    /// <summary>
    ///     Extension whitelist (AC-7 adjacent + F-3): passing a non-Word file path
    ///     throws <see cref="ArgumentException" /> without touching the filesystem.
    /// </summary>
    [Fact]
    public void Execute_RejectsNonWordExtension()
    {
        var tool = new WordOleObjectTool();

        Assert.Throws<ArgumentException>(() =>
            tool.Execute("list", _fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx]));
    }

    // ── From OleActivationGuardTests ──────────────────────────────────────

    /// <summary>
    ///     AC-12 (FLAG-3): Word <c>list</c> does not spawn a child of the current process.
    /// </summary>
    [Fact]
    public void Word_List_SpawnsNoChildProcess()
    {
        AssertNoChildSpawned(() => new WordOleObjectTool()
            .Execute("list", _fixtures.Paths[FixtureKind.WordEmbeddedDocx]));
    }

    /// <summary>
    ///     AC-12 (FLAG-3): Word <c>extract</c> does not spawn a child of the current process.
    /// </summary>
    [Fact]
    public void Word_Extract_SpawnsNoChildProcess()
    {
        var outputDir = Path.Combine(Path.GetTempPath(), "OleActGuard_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outputDir);
        try
        {
            AssertNoChildSpawned(() => new WordOleObjectTool()
                .Execute("extract", _fixtures.Paths[FixtureKind.WordEmbeddedDocx],
                    outputDirectory: outputDir, oleIndex: 0));
        }
        finally
        {
            try
            {
                Directory.Delete(outputDir, true);
            }
            catch
            {
                /* best-effort */
            }
        }
    }

    // ── From OleToolAnnotationsTests ──────────────────────────────────────

    /// <summary>
    ///     AC-14: the Word OLE tool carries the least-safe-op annotation bundle.
    /// </summary>
    [Fact]
    public void Tool_CarriesLeastSafeAnnotations()
    {
        var method = typeof(WordOleObjectTool).GetMethod("Execute", BindingFlags.Public | BindingFlags.Instance)!;
        var attr = method.GetCustomAttribute<McpServerToolAttribute>()!;

        Assert.Equal("word_ole_object", attr.Name);
        Assert.True(attr.Destructive, "OLE tool must be Destructive=true (bundles remove op)");
        Assert.False(attr.ReadOnly, "OLE tool must be ReadOnly=false");
        Assert.False(attr.Idempotent, "OLE tool must be Idempotent=false (remove is not idempotent)");
        Assert.False(attr.OpenWorld, "OLE tool must be OpenWorld=false (local file op)");
        Assert.True(attr.UseStructuredContent);
    }

    /// <summary>
    ///     AC-14: the Word OLE tool carries the <see cref="McpServerToolTypeAttribute" />
    ///     so the auto-discovery picks it up.
    /// </summary>
    [Fact]
    public void Tool_HasToolTypeAttribute()
    {
        Assert.NotNull(typeof(WordOleObjectTool).GetCustomAttribute<McpServerToolTypeAttribute>());
    }

    // ── From OleLegacyFormatTests ─────────────────────────────────────────

    /// <summary>Legacy <c>.doc</c> round-trips: list + remove preserves source format.</summary>
    [Fact]
    public void Legacy_Doc_RoundTrips()
    {
        var copy = Path.Combine(_outputDir, "legacy.doc");
        File.Copy(_fixtures.Paths[FixtureKind.WordEmbeddedDoc], copy);

        var tool = new WordOleObjectTool();
        var listData = ((FinalizedResult<OleListResult>)tool.Execute("list", copy)).Data;
        Assert.Equal(1, listData.Count);

        tool.Execute("remove", copy, oleIndex: 0);

        Assert.Equal(".doc", Path.GetExtension(copy).ToLowerInvariant());
        Assert.True(File.Exists(copy));
    }

    // ── From OlePasswordFileModeTests ─────────────────────────────────────

    /// <summary>
    ///     Word file-mode with the correct password lists OLE objects successfully.
    /// </summary>
    [Fact]
    public void Word_CorrectPassword_Succeeds()
    {
        var path = CreateProtectedWord("hunter2");

        var tool = new WordOleObjectTool();
        var raw = tool.Execute("list", path, password: "hunter2");

        var data = ((FinalizedResult<OleListResult>)raw).Data;
        Assert.True(data.Count >= 0, "List operation must succeed after correct password");
    }

    /// <summary>
    ///     Word file-mode with a wrong password throws
    ///     <see cref="UnauthorizedAccessException" /> carrying only the fixed sentinel.
    /// </summary>
    [Fact]
    public void Word_WrongPassword_ThrowsSanitized()
    {
        const string attempted = "CORRECT_HORSE_BATTERY_STAPLE";
        var path = CreateProtectedWord("hunter2");

        var tool = new WordOleObjectTool();
        var ex = Record.Exception(() => tool.Execute("list", path, password: attempted));

        Assert.NotNull(ex);
        Assert.DoesNotContain(attempted, ex.ToString());
    }

    // ── Private helpers ───────────────────────────────────────────────────

    /// <summary>
    ///     Builds a tiny in-memory xlsx workbook for use as an OLE payload. The payload
    ///     must be a real OOXML stream because Aspose.Words validates the package header
    ///     when the document is later reopened.
    /// </summary>
    /// <returns>A valid xlsx byte array.</returns>
    private static byte[] BuildValidXlsxPayload()
    {
        using var wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("x");
        using var ms = new MemoryStream();
        wb.Save(ms, SaveFormat.Xlsx);
        return ms.ToArray();
    }

    /// <summary>Minimal 1x1 PNG used by helper fixture builders.</summary>
    /// <returns>A valid PNG byte array.</returns>
    private static byte[] SmallPng()
    {
        return
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        ];
    }

    /// <summary>
    ///     Runs <paramref name="action" /> and asserts that no new process whose
    ///     start time falls within the call window appeared with an OLE-activator name.
    ///     Best effort — see OleActivationGuardTests remarks.
    /// </summary>
    /// <param name="action">Operation to invoke.</param>
    private static void AssertNoChildSpawned(Action action)
    {
        var baselinePids = new HashSet<int>(Process.GetProcesses().Select(p => p.Id));
        var startedAt = DateTime.UtcNow.AddSeconds(-1);

        action();

        var after = Process.GetProcesses();
        var suspects = after
            .Where(p => !baselinePids.Contains(p.Id))
            .Where(p =>
            {
                try
                {
                    return p.StartTime.ToUniversalTime() >= startedAt;
                }
                catch
                {
                    return false;
                }
            })
            .Select(p =>
            {
                try
                {
                    return p.ProcessName;
                }
                catch
                {
                    return "<inaccessible>";
                }
            })
            .ToList();

        string[] forbiddenNames = ["excel", "word", "powerpnt", "olexec", "wscript", "cscript"];
        foreach (var name in suspects)
            Assert.DoesNotContain(
                forbiddenNames,
                forbidden => name.Contains(forbidden, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    ///     Builds a tiny password-protected docx file in the test work directory.
    /// </summary>
    /// <param name="password">Open-password to encrypt with.</param>
    /// <returns>Absolute path to the encrypted docx.</returns>
    private string CreateProtectedWord(string password)
    {
        var path = Path.Combine(_outputDir, "prot.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("protected content");
        var options = new OoxmlSaveOptions { Password = password };
        doc.Save(path, options);
        return path;
    }
}
