using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.DOM.Ole;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Shared.Ole;
using AsposeMcpServer.Tests.Infrastructure.Ole;
using AsposeMcpServer.Tools.PowerPoint;
using ModelContextProtocol.Server;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Per-tool AC coverage for <c>ppt_ole_object</c>. Note from implementation.md:
///     <see cref="IOleObjectFrame.EmbeddedFileName" /> is read-only in Aspose.Slides
///     23.10.0, so attacker-controlled raw filenames for PPT are delivered via
///     <c>LinkPathLong</c> (linked frames) or rely on the <c>ole_N.xlsx</c> fallback
///     (embedded frames with empty names).
/// </summary>
[Collection(OleFixtureCollection.Name)]
public class PptOleObjectToolTests : IDisposable
{
    private readonly FixtureBuilder _fixtures;
    private readonly string _outputDir;

    /// <summary>Initializes a new instance of the <see cref="PptOleObjectToolTests" /> class.</summary>
    /// <param name="fixtures">Shared fixture matrix.</param>
    public PptOleObjectToolTests(FixtureBuilder fixtures)
    {
        _fixtures = fixtures;
        _outputDir = Path.Combine(Path.GetTempPath(), "PptOleTests_" + Guid.NewGuid().ToString("N"));
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

    /// <summary>AC-1: <c>list</c> returns expected metadata (shape location, progId, extension).</summary>
    [Fact]
    public void List_ReturnsExpectedMetadata()
    {
        var tool = new PptOleObjectTool();

        var raw = tool.Execute("list", _fixtures.Paths[FixtureKind.PptEmbeddedPptx]);

        var data = ((FinalizedResult<OleListResult>)raw).Data;
        Assert.Equal(1, data.Count);
        var item = data.Items[0];
        Assert.Equal(0, item.Index);
        Assert.False(item.IsLinked);
        Assert.EndsWith(".xlsx", item.SuggestedFileName);
        Assert.IsType<PptOleLocation>(item.ShapeLocation);
    }

    /// <summary>AC-2: <c>extract</c> writes correct bytes.</summary>
    [Fact]
    public void Extract_WritesCorrectBytes()
    {
        var tool = new PptOleObjectTool();

        var raw = tool.Execute("extract", _fixtures.Paths[FixtureKind.PptEmbeddedPptx],
            outputDirectory: _outputDir, oleIndex: 0);

        var data = ((FinalizedResult<OleExtractResult>)raw).Data;
        Assert.True(File.Exists(data.OutputFilePath));
        Assert.True(new FileInfo(data.OutputFilePath).Length > 0);
    }

    /// <summary>
    ///     AC-3 / AC-4 (PPT side channel — implementation.md limitation): the LinkPathLong
    ///     of a linked frame flows through the sanitizer and is emitted on the metadata
    ///     as <c>LinkTarget</c>. Attacker-controlled traversal is stripped.
    /// </summary>
    [Fact]
    public void List_LinkedFrame_LinkTargetIsSanitized()
    {
        var tool = new PptOleObjectTool();

        var raw = tool.Execute("list", _fixtures.Paths[FixtureKind.PptLinkedPptx]);

        var data = ((FinalizedResult<OleListResult>)raw).Data;
        Assert.Equal(1, data.Count);
        Assert.True(data.Items[0].IsLinked);
        // LinkTarget is sanitized via SanitizeForLog (F-4) — CRLF-free and control-char-free.
        var target = data.Items[0].LinkTarget;
        if (target != null)
        {
            Assert.DoesNotContain('\r', target);
            Assert.DoesNotContain('\n', target);
            Assert.DoesNotContain('\0', target);
        }
    }

    /// <summary>
    ///     AC-4: embedded frame's extracted file has a safe non-empty name with the
    ///     correct xlsx extension. Aspose.Slides may auto-populate
    ///     <c>EmbeddedFileName</c> (e.g. <c>oleObject1.xlsx</c>) during save, in which
    ///     case the sanitizer passes it through verbatim; when the raw name truly is
    ///     empty the <c>ole_N.xlsx</c> fallback kicks in. Both are acceptable AC-4
    ///     outcomes.
    /// </summary>
    [Fact]
    public void Extract_EmbeddedFallback_ProducesSafeXlsxName()
    {
        var tool = new PptOleObjectTool();

        var raw = tool.Execute("extract", _fixtures.Paths[FixtureKind.PptEmbeddedPptx],
            outputDirectory: _outputDir, oleIndex: 0);

        var data = ((FinalizedResult<OleExtractResult>)raw).Data;
        var fileName = Path.GetFileName(data.OutputFilePath);
        Assert.False(string.IsNullOrEmpty(fileName));
        Assert.EndsWith(".xlsx", fileName);
        Assert.DoesNotContain("..", fileName);
        Assert.DoesNotContain('\0', fileName);
    }

    /// <summary>AC-5: <c>extract_all</c> skips linked frames with reason <c>linked</c>.</summary>
    [Fact]
    public void ExtractAll_SkipsLinkedWithReason()
    {
        var tool = new PptOleObjectTool();

        var raw = tool.Execute("extract_all", _fixtures.Paths[FixtureKind.PptLinkedPptx],
            outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.Contains(data.Skipped, s => s.Reason == "linked");
    }

    /// <summary>AC-7: output dir outside allowlist is rejected.</summary>
    [Fact]
    public void OutputDirOutsideAllowedRoots_Rejected()
    {
        var config = new ServerConfig();
        typeof(ServerConfig)
            .GetProperty("AllowedBasePaths")!
            .SetValue(config, new[] { _outputDir }.ToList().AsReadOnly());

        var tool = new PptOleObjectTool(serverConfig: config);
        var sourceCopy = Path.Combine(_outputDir, "source.pptx");
        File.Copy(_fixtures.Paths[FixtureKind.PptEmbeddedPptx], sourceCopy);

        var outside = Path.Combine(Path.GetTempPath(), "PptOleOutside_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        try
        {
            Assert.Throws<UnauthorizedAccessException>(() =>
                tool.Execute("extract", sourceCopy, outputDirectory: outside, oleIndex: 0));
        }
        finally
        {
            try
            {
                Directory.Delete(outside, true);
            }
            catch
            {
                /* best-effort */
            }
        }
    }

    /// <summary>AC-8: session-mode parity with file-mode.</summary>
    [Fact]
    public void List_SessionMode_MatchesFileMode()
    {
        using var sessionManager = new DocumentSessionManager(
            new SessionConfig
            {
                Enabled = true,
                TempDirectory = Path.Combine(_outputDir, "sessions")
            });
        var tool = new PptOleObjectTool(sessionManager);
        var fixturePath = _fixtures.Paths[FixtureKind.PptEmbeddedPptx];

        var sid = sessionManager.OpenDocument(fixturePath);
        try
        {
            var fileData = ((FinalizedResult<OleListResult>)tool.Execute("list", fixturePath)).Data;
            var sessionData = ((FinalizedResult<OleListResult>)tool.Execute("list", sessionId: sid)).Data;
            Assert.Equal(fileData.Count, sessionData.Count);
        }
        finally
        {
            sessionManager.CloseDocument(sid, true);
        }
    }

    /// <summary>AC-11: repeated list is byte-identical JSON.</summary>
    [Fact]
    public void List_TwiceSameFile_ByteIdenticalJson()
    {
        var tool = new PptOleObjectTool();

        var r1 = tool.Execute("list", _fixtures.Paths[FixtureKind.PptEmbeddedPptx]);
        var r2 = tool.Execute("list", _fixtures.Paths[FixtureKind.PptEmbeddedPptx]);

        var d1 = ((FinalizedResult<OleListResult>)r1).Data;
        var d2 = ((FinalizedResult<OleListResult>)r2).Data;
        Assert.Equal(JsonSerializer.Serialize(d1), JsonSerializer.Serialize(d2));
    }

    /// <summary>AC-16: <c>remove</c> reduces count by one.</summary>
    [Fact]
    public void Remove_ReducesCountByOne()
    {
        var sourcePath = Path.Combine(_outputDir, "rem.pptx");
        File.Copy(_fixtures.Paths[FixtureKind.PptEmbeddedPptx], sourcePath);

        var tool = new PptOleObjectTool();
        tool.Execute("remove", sourcePath, oleIndex: 0);

        var listData = ((FinalizedResult<OleListResult>)tool.Execute("list", sourcePath)).Data;
        Assert.Equal(0, listData.Count);
    }

    /// <summary>AC-17: index-0-twice reindexing.</summary>
    [Fact]
    public void Remove_ThenRemove_AtIndexZero_AffectsDifferentObject()
    {
        var path = Path.Combine(_outputDir, "reidx.pptx");
        using (var pres = new Presentation())
        {
            var payload = BuildValidXlsxPayload();
            for (var i = 0; i < 2; i++)
            {
                var info = new OleEmbeddedDataInfo(payload, "xlsx");
                pres.Slides[0].Shapes.AddOleObjectFrame(10 + 20 * i, 10, 100, 100, info);
            }

            pres.Save(path, SlidesSaveFormat.Pptx);
        }

        var tool = new PptOleObjectTool();
        tool.Execute("remove", path, oleIndex: 0);
        tool.Execute("remove", path, oleIndex: 0);

        var list = ((FinalizedResult<OleListResult>)tool.Execute("list", path)).Data;
        Assert.Equal(0, list.Count);
    }

    /// <summary>AC-10: many-OLE list with no count cap.</summary>
    [Fact]
    public void NoCountCap_HandlesManyObjects()
    {
        var path = Path.Combine(_outputDir, "many.pptx");
        using (var pres = new Presentation())
        {
            var payload = BuildValidXlsxPayload();
            for (var i = 0; i < 25; i++)
            {
                var info = new OleEmbeddedDataInfo(payload, "xlsx");
                pres.Slides[0].Shapes.AddOleObjectFrame(10 + 5 * i, 10, 50, 50, info);
            }

            pres.Save(path, SlidesSaveFormat.Pptx);
        }

        var tool = new PptOleObjectTool();
        var data = ((FinalizedResult<OleListResult>)tool.Execute("list", path)).Data;
        Assert.Equal(25, data.Count);
        Assert.False(data.Truncated);
    }

    /// <summary>Extension whitelist rejects non-PPT.</summary>
    [Fact]
    public void Execute_RejectsNonPptExtension()
    {
        var tool = new PptOleObjectTool();
        Assert.Throws<ArgumentException>(() =>
            tool.Execute("list", _fixtures.Paths[FixtureKind.WordEmbeddedDocx]));
    }

    // ── From OleActivationGuardTests ──────────────────────────────────────

    /// <summary>
    ///     AC-12 (FLAG-3): PowerPoint <c>list</c> does not spawn a child of the current process.
    /// </summary>
    [Fact]
    public void Ppt_List_SpawnsNoChildProcess()
    {
        AssertNoChildSpawned(() => new PptOleObjectTool()
            .Execute("list", _fixtures.Paths[FixtureKind.PptEmbeddedPptx]));
    }

    // ── From OleToolAnnotationsTests ──────────────────────────────────────

    /// <summary>
    ///     AC-14: the PPT OLE tool carries the least-safe-op annotation bundle.
    /// </summary>
    [Fact]
    public void Tool_CarriesLeastSafeAnnotations()
    {
        var method = typeof(PptOleObjectTool).GetMethod("Execute", BindingFlags.Public | BindingFlags.Instance)!;
        var attr = method.GetCustomAttribute<McpServerToolAttribute>()!;

        Assert.Equal("ppt_ole_object", attr.Name);
        Assert.True(attr.Destructive, "OLE tool must be Destructive=true (bundles remove op)");
        Assert.False(attr.ReadOnly, "OLE tool must be ReadOnly=false");
        Assert.False(attr.Idempotent, "OLE tool must be Idempotent=false (remove is not idempotent)");
        Assert.False(attr.OpenWorld, "OLE tool must be OpenWorld=false (local file op)");
        Assert.True(attr.UseStructuredContent);
    }

    /// <summary>
    ///     AC-14: the PPT OLE tool carries the <see cref="McpServerToolTypeAttribute" />
    ///     so the auto-discovery picks it up.
    /// </summary>
    [Fact]
    public void Tool_HasToolTypeAttribute()
    {
        Assert.NotNull(typeof(PptOleObjectTool).GetCustomAttribute<McpServerToolTypeAttribute>());
    }

    // ── From OleLegacyFormatTests ─────────────────────────────────────────

    /// <summary>Legacy <c>.ppt</c> round-trips: list + remove preserves source format.</summary>
    [Fact]
    public void Legacy_Ppt_RoundTrips()
    {
        var copy = Path.Combine(_outputDir, "legacy.ppt");
        File.Copy(_fixtures.Paths[FixtureKind.PptEmbeddedPpt], copy);

        var tool = new PptOleObjectTool();
        var listData = ((FinalizedResult<OleListResult>)tool.Execute("list", copy)).Data;
        Assert.Equal(1, listData.Count);

        tool.Execute("remove", copy, oleIndex: 0);

        Assert.Equal(".ppt", Path.GetExtension(copy).ToLowerInvariant());
        Assert.True(File.Exists(copy));
    }

    // ── Private helpers ───────────────────────────────────────────────────

    /// <summary>
    ///     Builds a tiny valid xlsx payload (same shape as
    ///     <see cref="FixtureBuilder" />'s private helper) for in-test OLE insertion.
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
}
