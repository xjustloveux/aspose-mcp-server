using Aspose.Slides;
using Aspose.Slides.DOM.Ole;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Shared.Ole;
using AsposeMcpServer.Tests.Infrastructure.Ole;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Covers TEST-10 for PowerPoint. The cumulative byte cap on <c>extract_all</c> was only tested
///     through Excel; PowerPoint carries its own nested slide/shape loop with its own stop
///     condition, so nothing forced it to agree with the Excel behaviour.
/// </summary>
[Collection(OleFixtureCollection.Name)]
public class PptOleExtractAllSizeCapTests : IDisposable
{
    private readonly string _outputDir;
    private readonly string _presentationPath;

    /// <summary>
    ///     Builds a presentation with three embedded OLE frames of about 2 KiB each, spread over
    ///     two slides so the outer loop is exercised as well.
    /// </summary>
    /// <param name="_">Shared fixture collection; this test authors its own presentation.</param>
    // ReSharper disable once UnusedParameter.Local — xUnit injects the collection fixture
    public PptOleExtractAllSizeCapTests(FixtureBuilder _)
    {
        _outputDir = Path.Combine(Path.GetTempPath(), "PptOleSizeCap_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_outputDir);
        _presentationPath = Path.Combine(_outputDir, "cap.pptx");

        var payload = new byte[2048];
        Random.Shared.NextBytes(payload);

        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.Slides[0].LayoutSlide);

        presentation.Slides[0].Shapes.AddOleObjectFrame(10, 10, 200, 150,
            new OleEmbeddedDataInfo(payload, "bin"));
        presentation.Slides[0].Shapes.AddOleObjectFrame(10, 180, 200, 150,
            new OleEmbeddedDataInfo(payload, "bin"));
        presentation.Slides[1].Shapes.AddOleObjectFrame(10, 10, 200, 150,
            new OleEmbeddedDataInfo(payload, "bin"));

        presentation.Save(_presentationPath, SaveFormat.Pptx);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        try
        {
            Directory.Delete(_outputDir, true);
        }
        catch (IOException)
        {
            /* best-effort */
        }

        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ExtractAll_ExceedsCap_MarksTruncated()
    {
        var tool = new PptOleObjectTool(serverConfig: BuildConfigWithCap(4096));

        var raw = tool.Execute("extract_all", _presentationPath, outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.True(data.Truncated);
        Assert.Contains(data.Skipped, s => s.Reason == "cumulative-size-cap-exceeded");
        Assert.True(data.Extracted <= 2);
        OleExtractAllDiskAssertions.AssertDiskAgreesWithReport(
            _outputDir, _presentationPath, data, 4096);
    }

    [Fact]
    public void ExtractAll_UnderCap_CompletesEveryObject()
    {
        var tool = new PptOleObjectTool(serverConfig: BuildConfigWithCap(10L * 1024 * 1024 * 1024));

        var raw = tool.Execute("extract_all", _presentationPath, outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.False(data.Truncated);
        Assert.DoesNotContain(data.Skipped, s => s.Reason == "cumulative-size-cap-exceeded");
        Assert.Equal(3, data.Extracted);
        OleExtractAllDiskAssertions.AssertDiskAgreesWithReport(
            _outputDir, _presentationPath, data, 10L * 1024 * 1024 * 1024);
    }

    /// <summary>
    ///     Builds a <see cref="ServerConfig" /> with the cap set through reflection; the property
    ///     has a private setter so a test cannot assign it after construction.
    /// </summary>
    /// <param name="cap">Desired byte cap.</param>
    /// <returns>A config carrying that cap.</returns>
    private static ServerConfig BuildConfigWithCap(long cap)
    {
        var config = new ServerConfig();
        typeof(ServerConfig).GetProperty(nameof(ServerConfig.MaxExtractAllBytes))!.SetValue(config, cap);
        return config;
    }
}
