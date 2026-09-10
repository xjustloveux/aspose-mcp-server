using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Shared.Ole;
using AsposeMcpServer.Tests.Infrastructure.Ole;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Covers TEST-10 for Word. The cumulative byte cap on <c>extract_all</c> was only tested
///     through Excel on the grounds that the three handlers share the behaviour "by construction",
///     but each family carries its own copy of the loop, its own size computation and its own stop
///     condition, so nothing forced Word's copy to agree.
/// </summary>
[Collection(OleFixtureCollection.Name)]
public class WordOleExtractAllSizeCapTests : IDisposable
{
    private readonly string _documentPath;
    private readonly string _outputDir;

    /// <summary>
    ///     Builds a document with three embedded OLE packages of about 2 KiB each so the cap can be
    ///     placed between them.
    /// </summary>
    /// <param name="_">Shared fixture collection; this test authors its own document.</param>
    // ReSharper disable once UnusedParameter.Local — xUnit injects the collection fixture
    public WordOleExtractAllSizeCapTests(FixtureBuilder _)
    {
        _outputDir = Path.Combine(Path.GetTempPath(), "WordOleSizeCap_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_outputDir);
        _documentPath = Path.Combine(_outputDir, "cap.docx");

        var payload = new byte[2048];
        Random.Shared.NextBytes(payload);

        var document = new Document();
        var builder = new DocumentBuilder(document);
        for (var i = 0; i < 3; i++)
        {
            using var payloadStream = new MemoryStream(payload);
            using var imageStream = new MemoryStream(SubstitutePng);
            builder.InsertOleObject(payloadStream, "Package", false, imageStream);
            builder.Writeln();
        }

        var shapes = document.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToList();
        for (var i = 0; i < shapes.Count; i++)
            if (shapes[i].OleFormat?.OlePackage != null)
                shapes[i].OleFormat!.OlePackage!.FileName = $"payload_{i}.bin";

        document.Save(_documentPath, SaveFormat.Docx);
    }

    /// <summary>Smallest valid PNG, used as the OLE rendering substitute Aspose.Words requires.</summary>
    private static byte[] SubstitutePng =>
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
        var tool = new WordOleObjectTool(serverConfig: BuildConfigWithCap(4096));

        var raw = tool.Execute("extract_all", _documentPath, outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.True(data.Truncated);
        Assert.Contains(data.Skipped, s => s.Reason == "cumulative-size-cap-exceeded");
        Assert.True(data.Extracted <= 2);
        OleExtractAllDiskAssertions.AssertDiskAgreesWithReport(
            _outputDir, _documentPath, data, 4096);
    }

    [Fact]
    public void ExtractAll_UnderCap_CompletesEveryObject()
    {
        var tool = new WordOleObjectTool(serverConfig: BuildConfigWithCap(10L * 1024 * 1024 * 1024));

        var raw = tool.Execute("extract_all", _documentPath, outputDirectory: _outputDir);

        var data = ((FinalizedResult<OleExtractAllResult>)raw).Data;
        Assert.False(data.Truncated);
        Assert.DoesNotContain(data.Skipped, s => s.Reason == "cumulative-size-cap-exceeded");
        Assert.Equal(3, data.Extracted);
        OleExtractAllDiskAssertions.AssertDiskAgreesWithReport(
            _outputDir, _documentPath, data, 10L * 1024 * 1024 * 1024);
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
