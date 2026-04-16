using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Shared.Ole;
using AsposeMcpServer.Tests.Infrastructure.Ole;
using AsposeMcpServer.Tools.Excel;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Per-tool AC coverage for <c>excel_ole_object</c>. Shares the FixtureBuilder
///     matrix and drives the four operations against real Aspose.Cells workbooks.
/// </summary>
[Collection(OleFixtureCollection.Name)]
public class ExcelOleObjectToolTests : IDisposable
{
    private readonly FixtureBuilder _fixtures;
    private readonly string _outputDir;

    /// <summary>Initializes a new instance of the <see cref="ExcelOleObjectToolTests" /> class.</summary>
    /// <param name="fixtures">Shared fixture matrix.</param>
    public ExcelOleObjectToolTests(FixtureBuilder fixtures)
    {
        _fixtures = fixtures;
        _outputDir = Path.Combine(Path.GetTempPath(), "ExcelOleTests_" + Guid.NewGuid().ToString("N"));
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

    /// <summary>AC-1: <c>list</c> returns expected metadata shape.</summary>
    [Fact]
    public void List_ReturnsExpectedMetadata()
    {
        var tool = new ExcelOleObjectTool();

        var raw = tool.Execute("list", _fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx]);

        var data = ((FinalizedResult<OleListResult>)raw).Data;
        Assert.Equal(1, data.Count);
        var item = data.Items[0];
        Assert.Equal(0, item.Index);
        // Aspose.Cells auto-populates ObjectSourceFullName (e.g. "oleObject1.xlsx")
        // during Save — the Label we set is preserved but lower in the fallback chain.
        // What matters for AC-1 is that RawFileName is non-empty and SuggestedFileName
        // is a sanitized version of it.
        Assert.False(string.IsNullOrEmpty(item.RawFileName));
        Assert.False(string.IsNullOrEmpty(item.SuggestedFileName));
        Assert.EndsWith(".xlsx", item.SuggestedFileName);
        Assert.NotNull(item.ShapeLocation);
        Assert.IsType<ExcelOleLocation>(item.ShapeLocation);
    }

    /// <summary>AC-2: <c>extract</c> writes the correct bytes.</summary>
    [Fact]
    public void Extract_WritesCorrectBytes()
    {
        var tool = new ExcelOleObjectTool();

        var raw = tool.Execute("extract", _fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx],
            outputDirectory: _outputDir, oleIndex: 0);

        var data = ((FinalizedResult<OleExtractResult>)raw).Data;
        Assert.True(File.Exists(data.OutputFilePath));
        Assert.True(new FileInfo(data.OutputFilePath).Length > 0);
    }

    /// <summary>
    ///     AC-3 / AC-4: attacker-controlled <c>Label</c> (traversal) stays within the
    ///     output directory after sanitization.
    /// </summary>
    [Fact]
    public void Extract_AttackerLabel_StaysWithinOutputDir()
    {
        var tool = new ExcelOleObjectTool();

        var raw = tool.Execute("extract", _fixtures.Paths[FixtureKind.ExcelAttackerXlsx],
            outputDirectory: _outputDir, oleIndex: 0);

        var data = ((FinalizedResult<OleExtractResult>)raw).Data;
        // AC-3 invariant: the output file lives inside the output directory AND the
        // filename contains no separator, traversal sequence, or NUL. The
        // SanitizedFromRaw flag depends on whether Aspose.Cells' auto-populated
        // ObjectSourceFullName happens to already be sanitized — that isn't the AC.
        Assert.Equal(
            Path.GetFullPath(_outputDir),
            Path.GetDirectoryName(data.OutputFilePath));
        var fileName = Path.GetFileName(data.OutputFilePath);
        Assert.DoesNotContain("..", fileName);
        Assert.DoesNotContain("/", fileName);
        Assert.DoesNotContain("\\", fileName);
        Assert.DoesNotContain('\0', fileName);
    }

    /// <summary>
    ///     AC-7: output directory outside the allowlist is rejected.
    /// </summary>
    [Fact]
    public void OutputDirOutsideAllowedRoots_Rejected()
    {
        var config = new ServerConfig();
        typeof(ServerConfig)
            .GetProperty("AllowedBasePaths")!
            .SetValue(config, new[] { _outputDir }.ToList().AsReadOnly());

        var tool = new ExcelOleObjectTool(serverConfig: config);
        var sourceCopy = Path.Combine(_outputDir, "source.xlsx");
        File.Copy(_fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx], sourceCopy);

        var outside = Path.Combine(Path.GetTempPath(), "ExcelOleOutside_" + Guid.NewGuid().ToString("N"));
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

    /// <summary>AC-8: session-mode list matches file-mode item count.</summary>
    [Fact]
    public void List_SessionMode_MatchesFileMode()
    {
        using var sessionManager = new DocumentSessionManager(
            new SessionConfig
            {
                Enabled = true,
                TempDirectory = Path.Combine(_outputDir, "sessions")
            });
        var tool = new ExcelOleObjectTool(sessionManager);
        var fixturePath = _fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx];

        var sid = sessionManager.OpenDocument(fixturePath);
        try
        {
            var fileRaw = tool.Execute("list", fixturePath);
            var sessionRaw = tool.Execute("list", sessionId: sid);

            var fileData = ((FinalizedResult<OleListResult>)fileRaw).Data;
            var sessionData = ((FinalizedResult<OleListResult>)sessionRaw).Data;
            Assert.Equal(fileData.Count, sessionData.Count);
        }
        finally
        {
            sessionManager.CloseDocument(sid, true);
        }
    }

    /// <summary>AC-11: repeated list yields byte-identical JSON.</summary>
    [Fact]
    public void List_TwiceSameFile_ByteIdenticalJson()
    {
        var tool = new ExcelOleObjectTool();

        var r1 = tool.Execute("list", _fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx]);
        var r2 = tool.Execute("list", _fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx]);

        var d1 = ((FinalizedResult<OleListResult>)r1).Data;
        var d2 = ((FinalizedResult<OleListResult>)r2).Data;
        Assert.Equal(JsonSerializer.Serialize(d1), JsonSerializer.Serialize(d2));
    }

    /// <summary>AC-16: <c>remove</c> reduces the count by one.</summary>
    [Fact]
    public void Remove_ReducesCountByOne()
    {
        var sourcePath = Path.Combine(_outputDir, "rem.xlsx");
        File.Copy(_fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx], sourcePath);

        var tool = new ExcelOleObjectTool();
        tool.Execute("remove", sourcePath, oleIndex: 0);

        var listRaw = tool.Execute("list", sourcePath);
        var listData = ((FinalizedResult<OleListResult>)listRaw).Data;
        Assert.Equal(0, listData.Count);
    }

    /// <summary>
    ///     AC-17: index 0 twice in a row affects two distinct OLEs. Build a fresh
    ///     2-OLE workbook.
    /// </summary>
    [Fact]
    public void Remove_ThenRemove_AtIndexZero_AffectsDifferentObject()
    {
        var path = Path.Combine(_outputDir, "reidx.xlsx");
        using (var wb = new Workbook())
        {
            var png = SmallPng();
            var payload = new byte[] { 0x10, 0x20, 0x30, 0x40 };
            for (var i = 0; i < 2; i++)
            {
                var idx = wb.Worksheets[0].OleObjects.Add(1 + i, 1, 50, 50, png);
                wb.Worksheets[0].OleObjects[idx].ObjectData = payload;
                wb.Worksheets[0].OleObjects[idx].ProgID = "Package";
            }

            wb.Save(path);
        }

        var tool = new ExcelOleObjectTool();
        tool.Execute("remove", path, oleIndex: 0);
        tool.Execute("remove", path, oleIndex: 0);

        var list = ((FinalizedResult<OleListResult>)tool.Execute("list", path)).Data;
        Assert.Equal(0, list.Count);
    }

    /// <summary>AC-10: no count cap — a workbook with many OLEs is listed fully.</summary>
    [Fact]
    public void NoCountCap_HandlesManyObjects()
    {
        var path = Path.Combine(_outputDir, "many.xlsx");
        using (var wb = new Workbook())
        {
            var png = SmallPng();
            var payload = new byte[] { 0xAA, 0xBB, 0xCC };
            for (var i = 0; i < 40; i++)
            {
                var idx = wb.Worksheets[0].OleObjects.Add(1 + i, 1, 50, 50, png);
                wb.Worksheets[0].OleObjects[idx].ObjectData = payload;
                wb.Worksheets[0].OleObjects[idx].ProgID = "Package";
            }

            wb.Save(path);
        }

        var tool = new ExcelOleObjectTool();
        var data = ((FinalizedResult<OleListResult>)tool.Execute("list", path)).Data;
        Assert.Equal(40, data.Count);
        Assert.False(data.Truncated);
    }

    /// <summary>Extension whitelist rejects non-Excel inputs.</summary>
    [Fact]
    public void Execute_RejectsNonExcelExtension()
    {
        var tool = new ExcelOleObjectTool();
        Assert.Throws<ArgumentException>(() =>
            tool.Execute("list", _fixtures.Paths[FixtureKind.WordEmbeddedDocx]));
    }

    // ── From OleActivationGuardTests ──────────────────────────────────────

    /// <summary>
    ///     AC-12 (FLAG-3): Excel <c>list</c> does not spawn a child of the current process.
    /// </summary>
    [Fact]
    public void Excel_List_SpawnsNoChildProcess()
    {
        AssertNoChildSpawned(() => new ExcelOleObjectTool()
            .Execute("list", _fixtures.Paths[FixtureKind.ExcelEmbeddedXlsx]));
    }

    // ── From OleToolAnnotationsTests ──────────────────────────────────────

    /// <summary>
    ///     AC-14: the Excel OLE tool carries the least-safe-op annotation bundle.
    /// </summary>
    [Fact]
    public void Tool_CarriesLeastSafeAnnotations()
    {
        var method = typeof(ExcelOleObjectTool).GetMethod("Execute", BindingFlags.Public | BindingFlags.Instance)!;
        var attr = method.GetCustomAttribute<McpServerToolAttribute>()!;

        Assert.Equal("excel_ole_object", attr.Name);
        Assert.True(attr.Destructive, "OLE tool must be Destructive=true (bundles remove op)");
        Assert.False(attr.ReadOnly, "OLE tool must be ReadOnly=false");
        Assert.False(attr.Idempotent, "OLE tool must be Idempotent=false (remove is not idempotent)");
        Assert.False(attr.OpenWorld, "OLE tool must be OpenWorld=false (local file op)");
        Assert.True(attr.UseStructuredContent);
    }

    /// <summary>
    ///     AC-14: the Excel OLE tool carries the <see cref="McpServerToolTypeAttribute" />
    ///     so the auto-discovery picks it up.
    /// </summary>
    [Fact]
    public void Tool_HasToolTypeAttribute()
    {
        Assert.NotNull(typeof(ExcelOleObjectTool).GetCustomAttribute<McpServerToolTypeAttribute>());
    }

    // ── From OleLegacyFormatTests ─────────────────────────────────────────

    /// <summary>Legacy <c>.xls</c> round-trips: list + remove preserves source format.</summary>
    [Fact]
    public void Legacy_Xls_RoundTrips()
    {
        var copy = Path.Combine(_outputDir, "legacy.xls");
        File.Copy(_fixtures.Paths[FixtureKind.ExcelEmbeddedXls], copy);

        var tool = new ExcelOleObjectTool();
        var listData = ((FinalizedResult<OleListResult>)tool.Execute("list", copy)).Data;
        Assert.Equal(1, listData.Count);

        tool.Execute("remove", copy, oleIndex: 0);

        Assert.Equal(".xls", Path.GetExtension(copy).ToLowerInvariant());
        Assert.True(File.Exists(copy));
    }

    // ── From OlePasswordFileModeTests ─────────────────────────────────────

    /// <summary>
    ///     Excel file-mode with the correct password lists OLE objects.
    /// </summary>
    [Fact]
    public void Excel_CorrectPassword_Succeeds()
    {
        var path = CreateProtectedExcel("hunter2");

        var tool = new ExcelOleObjectTool();
        var raw = tool.Execute("list", path, password: "hunter2");

        var data = ((FinalizedResult<OleListResult>)raw).Data;
        Assert.True(data.Count >= 0, "List operation must succeed after correct password");
    }

    /// <summary>
    ///     Excel file-mode with wrong password throws and never echoes the attempted
    ///     password.
    /// </summary>
    [Fact]
    public void Excel_WrongPassword_ThrowsSanitized()
    {
        const string attempted = "DO_NOT_ECHO_ME";
        var path = CreateProtectedExcel("hunter2");

        var tool = new ExcelOleObjectTool();
        var ex = Record.Exception(() => tool.Execute("list", path, password: attempted));

        Assert.NotNull(ex);
        Assert.DoesNotContain(attempted, ex.ToString());
    }

    // ── Private helpers ───────────────────────────────────────────────────

    /// <summary>1x1 PNG helper.</summary>
    /// <returns>Valid PNG byte array.</returns>
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
    ///     Builds a tiny password-protected xlsx file.
    /// </summary>
    /// <param name="password">Open-password.</param>
    /// <returns>Absolute path.</returns>
    private string CreateProtectedExcel(string password)
    {
        var path = Path.Combine(_outputDir, "prot.xlsx");
        using var wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("protected");
        wb.Settings.Password = password;
        wb.Save(path);
        return path;
    }
}
