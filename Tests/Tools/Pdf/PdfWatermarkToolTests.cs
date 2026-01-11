using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfWatermarkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfWatermarkToolTests : PdfTestBase
{
    private readonly PdfWatermarkTool _tool;

    public PdfWatermarkToolTests()
    {
        _tool = new PdfWatermarkTool(SessionManager);
    }

    private string CreatePdfDocument(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} Content"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [SkippableFact]
    public void Add_WithText_ShouldAddWatermark()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Watermark text extraction has limitations in evaluation mode");
        const string watermarkText = "Confidential";
        var pdfPath = CreatePdfDocument("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 1 page(s)", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [SkippableFact]
    public void Add_WithMultiplePages_ShouldApplyToAllPages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Watermark text extraction has limitations in evaluation mode");
        const string watermarkText = "All Pages";
        var pdfPath = CreatePdfDocument("test_multi.pdf", 3);
        var outputPath = CreateTestFilePath("test_multi_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 3 page(s)", result);
    }

    [Fact]
    public void Add_WithPageRange_ShouldApplyToSpecificPages()
    {
        const string watermarkText = "Range Pages";
        var pdfPath = CreatePdfDocument("test_hyphen.pdf", 4);
        var outputPath = CreateTestFilePath("test_hyphen_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, pageRange: "2-4");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 3 page(s)", result);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");

        var result = _tool.Execute(operation, text: "Watermark", path: pdfPath, outputPath: outputPath);

        Assert.StartsWith("Watermark added to 1 page(s)", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", text: "Test", path: pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("add", text: "Test", path: null, sessionId: null));
    }

    #endregion

    #region Session Management

    [SkippableFact]
    public void Add_WithSessionId_ShouldModifyInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Watermark text extraction has limitations in evaluation mode");
        const string watermarkText = "Confidential";
        var pdfPath = CreatePdfDocument("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", text: watermarkText, sessionId: sessionId);

        Assert.StartsWith("Watermark added to 1 page(s)", result);
        Assert.Contains(sessionId, result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);

        var textAbsorber = new TextAbsorber();
        document.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithSessionId_AndPageRange_ShouldApplyToSpecificPages()
    {
        const string watermarkText = "Selected";
        var pdfPath = CreatePdfDocument("test_session_range.pdf", 3);
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", text: watermarkText, sessionId: sessionId, pageRange: "1,3");

        Assert.StartsWith("Watermark added to 2 page(s)", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("add", text: "Test", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreatePdfDocument("test_path_watermark.pdf");
        var pdfPath2 = CreatePdfDocument("test_session_watermark.pdf", 3);
        var sessionId = OpenSession(pdfPath2);

        var result = _tool.Execute("add", text: "Test", path: pdfPath1, sessionId: sessionId);

        Assert.StartsWith("Watermark added to 3 page(s)", result);
    }

    #endregion
}
