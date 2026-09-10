using System.Reflection;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.DOM.Ole;
using AsposeMcpServer.Core;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;
using AsposeMcpServer.Tools.Ocr;
using AsposeMcpServer.Tools.Pdf;
using AsposeMcpServer.Tools.PowerPoint;
using AsposeMcpServer.Tools.Word;
using Microsoft.Extensions.DependencyInjection;
using SkiaSharp;
using WordsDocument = Aspose.Words.Document;
using WordsDocumentBuilder = Aspose.Words.DocumentBuilder;
using PdfDocument = Aspose.Pdf.Document;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools;

/// <summary>
///     Tool-layer integration guard for the <c>--allowed-path</c> allowlist (RB-01).
///     Handler-level tests such as <see cref="Handlers.AllowlistBypassHandlerTests" /> construct an
///     <see cref="AsposeMcpServer.Core.Handlers.OperationContext{T}" /> by hand and inject a
///     <see cref="ServerConfig" />, so they cannot detect a Tool that never passes the configured
///     config down to <see cref="AsposeMcpServer.Core.Session.DocumentContext{T}" /> or to the handler context.
///     These tests build each Tool through a real <see cref="IServiceProvider" /> exactly the way
///     the MCP host does (<see cref="ActivatorUtilities" />), so a missing constructor parameter or
///     a missing pass-through fails here.
/// </summary>
public class ToolAllowlistWiringTests : TestBase
{
    private readonly string _outsideDir;

    /// <summary>
    ///     Initializes the fixture with a directory that is deliberately outside the allowlist.
    /// </summary>
    public ToolAllowlistWiringTests()
    {
        _outsideDir = Path.Combine(Path.GetTempPath(), "AsposeAllowlistOutside_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_outsideDir);
    }

    /// <summary>
    ///     Builds a <see cref="ServerConfig" /> whose allowlist contains only <see cref="TestBase.TestDir" />.
    ///     The property setter is private, matching the pattern used by the handler-level allowlist tests.
    /// </summary>
    /// <returns>A configured <see cref="ServerConfig" />.</returns>
    private ServerConfig BuildAllowlistedConfig()
    {
        var cfg = new ServerConfig();
        var prop = typeof(ServerConfig).GetProperty(
            nameof(ServerConfig.AllowedBasePaths),
            BindingFlags.Instance | BindingFlags.Public);
        prop!.SetValue(cfg, new List<string> { Path.GetFullPath(TestDir) });
        return cfg;
    }

    /// <summary>
    ///     Creates the Tool the same way the MCP host does: through dependency injection with the
    ///     server configuration registered as a singleton.
    /// </summary>
    /// <typeparam name="TTool">The tool type to construct.</typeparam>
    /// <returns>A tool instance with whatever services it declares resolved from the container.</returns>
    private TTool CreateTool<TTool>() where TTool : class
    {
        var services = new ServiceCollection();
        services.AddSingleton(BuildAllowlistedConfig());
        var provider = services.BuildServiceProvider();
        return ActivatorUtilities.CreateInstance<TTool>(provider);
    }

    /// <summary>
    ///     Creates an empty workbook at the given path.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <returns>The path that was written.</returns>
    private static string WriteWorkbook(string path)
    {
        using var workbook = new Workbook();
        workbook.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates a one-paragraph Word document at the given path.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <returns>The path that was written.</returns>
    private static string WriteWordDocument(string path)
    {
        var doc = new WordsDocument();
        var builder = new WordsDocumentBuilder(doc);
        builder.Writeln("allowlist fixture");
        doc.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates a one-slide presentation at the given path.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <returns>The path that was written.</returns>
    private static string WritePresentation(string path)
    {
        using var presentation = new Presentation();
        presentation.Save(path, SaveFormat.Pptx);
        return path;
    }

    /// <summary>
    ///     Creates a one-page PDF at the given path.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <returns>The path that was written.</returns>
    private static string WritePdf(string path)
    {
        using var pdf = new PdfDocument();
        pdf.Pages.Add();
        pdf.Save(path);
        return path;
    }

    /// <summary>
    ///     Builds the small xlsx payload embedded as the OLE object in the fixtures below.
    /// </summary>
    /// <returns>Bytes of a valid workbook.</returns>
    private static byte[] BuildOlePayload()
    {
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("ole allowlist fixture");
        using var stream = new MemoryStream();
        workbook.Save(stream, Aspose.Cells.SaveFormat.Xlsx);
        return stream.ToArray();
    }

    /// <summary>
    ///     Builds the substitute picture Aspose requires when embedding an OLE object.
    /// </summary>
    /// <returns>PNG bytes.</returns>
    private static byte[] BuildSubstitutePng()
    {
        using var bitmap = new SKBitmap(8, 8);
        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    /// <summary>
    ///     Creates a Word document carrying one embedded OLE object.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <returns>The path that was written.</returns>
    private static string WriteWordWithOle(string path)
    {
        var doc = new WordsDocument();
        var builder = new WordsDocumentBuilder(doc);
        using var payload = new MemoryStream(BuildOlePayload());
        using var picture = new MemoryStream(BuildSubstitutePng());
        builder.InsertOleObject(payload, "Excel.Sheet.12", false, picture);
        doc.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates a workbook carrying one embedded OLE object.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <returns>The path that was written.</returns>
    private static string WriteWorkbookWithOle(string path)
    {
        using var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var index = sheet.OleObjects.Add(1, 1, 100, 100, BuildSubstitutePng());
        var ole = sheet.OleObjects[index];
        ole.ProgID = "Excel.Sheet.12";
        ole.ObjectData = BuildOlePayload();
        workbook.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates a presentation carrying one embedded OLE object.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <returns>The path that was written.</returns>
    private static string WritePresentationWithOle(string path)
    {
        using var presentation = new Presentation();
        presentation.Slides[0].Shapes.AddOleObjectFrame(
            10, 10, 200, 150, new OleEmbeddedDataInfo(BuildOlePayload(), "xlsx"));
        presentation.Save(path, SaveFormat.Pptx);
        return path;
    }

    /// <summary>Path to a file placed outside the allowlist.</summary>
    /// <param name="fileName">File name to create under the outside directory.</param>
    /// <returns>The absolute path.</returns>
    private string Outside(string fileName)
    {
        return Path.Combine(_outsideDir, fileName);
    }

    [Fact]
    public void ExcelSheet_Get_OutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<ExcelSheetTool>();
        var path = WriteWorkbook(Outside("read.xlsx"));

        Assert.Throws<ArgumentException>(() => tool.Execute("get", path));
    }

    [Fact]
    public void ExcelSheet_OutputPathOutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<ExcelSheetTool>();
        var path = WriteWorkbook(CreateTestFilePath("source.xlsx"));
        var outputPath = Outside("written.xlsx");

        Assert.Throws<ArgumentException>(() =>
            tool.Execute("add", path, outputPath: outputPath, sheetName: "Added"));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void WordParagraph_Get_OutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<WordParagraphTool>();
        var path = WriteWordDocument(Outside("read.docx"));

        Assert.Throws<ArgumentException>(() => tool.Execute("get", path));
    }

    [Fact]
    public void WordParagraph_OutputPathOutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<WordParagraphTool>();
        var path = WriteWordDocument(CreateTestFilePath("source.docx"));
        var outputPath = Outside("written.docx");

        Assert.Throws<ArgumentException>(() =>
            tool.Execute("insert", path, outputPath: outputPath, text: "new paragraph"));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void PptSlide_Get_OutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<PptSlideTool>();
        var path = WritePresentation(Outside("read.pptx"));

        Assert.Throws<ArgumentException>(() => tool.Execute("get", path));
    }

    [Fact]
    public void PptSlide_OutputPathOutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<PptSlideTool>();
        var path = WritePresentation(CreateTestFilePath("source.pptx"));
        var outputPath = Outside("written.pptx");

        Assert.Throws<ArgumentException>(() => tool.Execute("add", path, outputPath: outputPath));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void PdfPage_Info_OutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<PdfPageTool>();
        var path = WritePdf(Outside("read.pdf"));

        Assert.Throws<ArgumentException>(() => tool.Execute("info", path));
    }

    [Fact]
    public void PdfPage_OutputPathOutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<PdfPageTool>();
        var path = WritePdf(CreateTestFilePath("source.pdf"));
        var outputPath = Outside("written.pdf");

        Assert.Throws<ArgumentException>(() => tool.Execute("add", path, outputPath: outputPath));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void WordProperties_Get_OutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<WordPropertiesTool>();
        var path = WriteWordDocument(Outside("props.docx"));

        Assert.Throws<ArgumentException>(() => tool.Execute("get", path));
    }

    [Fact]
    public void OcrRecognition_InputOutsideAllowlist_ShouldBeRejected()
    {
        // The recognition handlers read context.ServerConfig?.AllowedBasePaths ?? [], and an empty
        // allowlist means unrestricted, so a tool that never sets ServerConfig silently opts every
        // OCR operation out of --allowed-path (R2-S01). The allowlist check runs before the file is
        // opened, so this needs no image and no OCR licence.
        var tool = CreateTool<OcrRecognitionTool>();
        var path = Outside("scan.png");
        File.WriteAllBytes(path, [0x89, 0x50, 0x4E, 0x47]);

        Assert.Throws<ArgumentException>(() => tool.Execute("recognize", path));
    }

    [Fact]
    public void WordOle_RemoveOutputPathOutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<WordOleObjectTool>();
        var path = WriteWordWithOle(CreateTestFilePath("ole_source.docx"));
        var outputPath = Outside("ole_written.docx");

        Assert.Throws<ArgumentException>(() =>
            tool.Execute("remove", path, oleIndex: 0, outputPath: outputPath));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void ExcelOle_RemoveOutputPathOutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<ExcelOleObjectTool>();
        var path = WriteWorkbookWithOle(CreateTestFilePath("ole_source.xlsx"));
        var outputPath = Outside("ole_written.xlsx");

        Assert.Throws<ArgumentException>(() =>
            tool.Execute("remove", path, oleIndex: 0, outputPath: outputPath));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void PptOle_RemoveOutputPathOutsideAllowlist_ShouldBeRejected()
    {
        var tool = CreateTool<PptOleObjectTool>();
        var path = WritePresentationWithOle(CreateTestFilePath("ole_source.pptx"));
        var outputPath = Outside("ole_written.pptx");

        Assert.Throws<ArgumentException>(() =>
            tool.Execute("remove", path, oleIndex: 0, outputPath: outputPath));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void ExcelSheet_InsideAllowlist_ShouldSucceed()
    {
        var tool = CreateTool<ExcelSheetTool>();
        var path = WriteWorkbook(CreateTestFilePath("inside.xlsx"));
        var outputPath = CreateTestFilePath("inside_out.xlsx");

        tool.Execute("add", path, outputPath: outputPath, sheetName: "Added");

        Assert.True(File.Exists(outputPath));
    }

    /// <inheritdoc />
    public override void Dispose()
    {
        if (Directory.Exists(_outsideDir))
            try
            {
                Directory.Delete(_outsideDir, true);
            }
            catch (IOException)
            {
                // Best-effort cleanup of the fixture directory; a locked file must not fail the test run.
            }

        base.Dispose();
    }
}
