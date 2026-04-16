using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Shared.Ole;
using AsposeMcpServer.Tests.Infrastructure.Ole;
using AsposeMcpServer.Tools.Excel;
using CellsSaveFormat = Aspose.Cells.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     AC-22 coverage (F-8): when the cumulative bytes written by <c>extract_all</c>
///     exceed <see cref="ServerConfig.MaxExtractAllBytes" />, the operation must halt,
///     mark the remaining entries skipped with reason
///     <c>cumulative-size-cap-exceeded</c>, and set <c>truncated: true</c>. Uses Excel
///     because the helper is identical across the three handlers (F-8 "by construction"
///     per implementation.md).
/// </summary>
[Collection(OleFixtureCollection.Name)]
public class OleExtractAllSizeCapTests : IDisposable
{
    private readonly string _outputDir;
    private readonly string _workbookPath;

    /// <summary>
    ///     Builds a tiny Excel workbook with three OLE payloads of ~2 KiB each so the
    ///     handler can exercise the cap at a known threshold.
    /// </summary>
    /// <param name="_">Shared fixture builder (unused — we author the fixture inline).</param>
    // ReSharper disable once UnusedParameter.Local — xUnit injects the collection fixture; this test builds its own workbook inline
    public OleExtractAllSizeCapTests(FixtureBuilder _)
    {
        _outputDir = Path.Combine(Path.GetTempPath(), "OleSizeCap_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_outputDir);

        _workbookPath = Path.Combine(_outputDir, "cap.xlsx");
        using var wb = new Workbook();
        var sheet = wb.Worksheets[0];
        var payload = new byte[2048];
        Random.Shared.NextBytes(payload);
        var png = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        };
        for (var i = 0; i < 3; i++)
        {
            var idx = sheet.OleObjects.Add(1 + i, 1, 50, 50, png);
            var ole = sheet.OleObjects[idx];
            ole.ProgID = "Package";
            ole.ObjectData = payload;
            ole.Label = $"payload_{i}.bin";
        }

        wb.Save(_workbookPath, CellsSaveFormat.Xlsx);
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
    ///     With a 4 KiB cap and three 2 KiB payloads, the third OLE is skipped with
    ///     <c>cumulative-size-cap-exceeded</c> and <c>truncated</c> is <c>true</c>.
    /// </summary>
    [Fact]
    public void ExtractAll_ExceedsCap_MarksTruncated()
    {
        var config = BuildConfigWithCap(4096);
        var tool = new ExcelOleObjectTool(serverConfig: config);

        var raw = tool.Execute(
            "extract_all",
            _workbookPath,
            outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.True(data.Truncated);
        Assert.Contains(data.Skipped, s => s.Reason == "cumulative-size-cap-exceeded");
        Assert.True(data.Extracted <= 2);
    }

    /// <summary>
    ///     With the default 10 GiB cap, the same extract_all succeeds for all three
    ///     payloads and <c>truncated</c> remains <c>false</c> — proving the cap only
    ///     kicks in when the cumulative total actually exceeds the limit.
    /// </summary>
    [Fact]
    public void ExtractAll_UnderCap_CompletesAllThree()
    {
        var config = BuildConfigWithCap(10L * 1024 * 1024 * 1024);
        var tool = new ExcelOleObjectTool(serverConfig: config);

        var raw = tool.Execute(
            "extract_all",
            _workbookPath,
            outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.False(data.Truncated);
        Assert.DoesNotContain(data.Skipped, s => s.Reason == "cumulative-size-cap-exceeded");
        Assert.Equal(3, data.Extracted);
    }

    /// <summary>
    ///     Builds a <see cref="ServerConfig" /> with <see cref="ServerConfig.MaxExtractAllBytes" />
    ///     set via reflection — the property is init-only with a private setter so tests
    ///     cannot assign it after construction via public API.
    /// </summary>
    /// <param name="cap">Desired byte cap.</param>
    /// <returns>A config with the requested cap.</returns>
    private static ServerConfig BuildConfigWithCap(long cap)
    {
        var config = new ServerConfig();
        var prop = typeof(ServerConfig).GetProperty(nameof(ServerConfig.MaxExtractAllBytes))!;
        prop.SetValue(config, cap);
        return config;
    }
}
