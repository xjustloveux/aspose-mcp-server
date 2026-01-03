using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfRedactToolTests : PdfTestBase
{
    private readonly PdfRedactTool _tool;

    public PdfRedactToolTests()
    {
        _tool = new PdfRedactTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Text to redact"));
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void RedactArea_ShouldRedactArea()
    {
        var pdfPath = CreateTestPdf("test_redact_area.pdf");
        var outputPath = CreateTestFilePath("test_redact_area_output.pdf");
        _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50);
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void RedactWithColor_ShouldRedactWithColor()
    {
        var pdfPath = CreateTestPdf("test_redact_color.pdf");
        var outputPath = CreateTestFilePath("test_redact_color_output.pdf");
        _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50,
            fillColor: "255,0,0");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void RedactWithOverlayText_ShouldRedactWithText()
    {
        var pdfPath = CreateTestPdf("test_redact_overlay.pdf");
        var outputPath = CreateTestFilePath("test_redact_overlay_output.pdf");
        var result = _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50,
            overlayText: "[REDACTED]");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public void RedactWithColorAndOverlayText_ShouldApplyBoth()
    {
        var pdfPath = CreateTestPdf("test_redact_both.pdf");
        var outputPath = CreateTestFilePath("test_redact_both_output.pdf");
        var result = _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50,
            fillColor: "Red",
            overlayText: "CONFIDENTIAL");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public void Redact_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_redact_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            pdfPath,
            pageIndex: 99,
            x: 100,
            y: 100,
            width: 200,
            height: 50));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void Redact_WithColorName_ShouldParseColor()
    {
        var pdfPath = CreateTestPdf("test_redact_color_name.pdf");
        var outputPath = CreateTestFilePath("test_redact_color_name_output.pdf");
        var result = _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50,
            fillColor: "Blue");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public void Redact_WithDefaultOutput_ShouldOverwriteInput()
    {
        var pdfPath = CreateTestPdf("test_redact_default_output.pdf");
        var result = _tool.Execute(
            pdfPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50);
        Assert.True(File.Exists(pdfPath), "PDF should still exist");
        Assert.Contains("Redaction applied", result);
    }

    [Fact]
    public void RedactByText_ShouldFindAndRedactText()
    {
        var pdfPath = CreateTestPdf("test_redact_by_text.pdf");
        var outputPath = CreateTestFilePath("test_redact_by_text_output.pdf");
        var result = _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            textToRedact: "redact");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Redacted", result);
        Assert.Contains("occurrence", result);
    }

    [Fact]
    public void RedactByText_OnSpecificPage_ShouldRedactOnlyThatPage()
    {
        var pdfPath = CreateTestPdf("test_redact_text_page.pdf");
        var outputPath = CreateTestFilePath("test_redact_text_page_output.pdf");
        _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            textToRedact: "Text");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void RedactByText_CaseInsensitive_ShouldFindText()
    {
        var pdfPath = CreateTestPdf("test_redact_case_insensitive.pdf");
        var outputPath = CreateTestFilePath("test_redact_case_insensitive_output.pdf");
        _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            textToRedact: "TEXT",
            caseSensitive: false);
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void RedactByText_NotFound_ShouldReturnNoOccurrences()
    {
        var pdfPath = CreateTestPdf("test_redact_not_found.pdf");
        var result = _tool.Execute(
            pdfPath,
            textToRedact: "nonexistent_text_12345");
        Assert.Contains("No occurrences", result);
        Assert.Contains("No redactions applied", result);
    }

    [Fact]
    public void RedactByText_WithOverlayText_ShouldApplyOverlay()
    {
        var pdfPath = CreateTestPdf("test_redact_text_overlay.pdf");
        var outputPath = CreateTestFilePath("test_redact_text_overlay_output.pdf");
        _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            textToRedact: "Text",
            overlayText: "[CLASSIFIED]");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void RedactByText_WithColor_ShouldApplyColor()
    {
        var pdfPath = CreateTestPdf("test_redact_text_color.pdf");
        var outputPath = CreateTestFilePath("test_redact_text_color_output.pdf");
        _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            textToRedact: "redact",
            fillColor: "Red");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void RedactByText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_redact_text_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            pdfPath,
            pageIndex: 99,
            textToRedact: "Text"));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    private string CreateMultiPagePdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        for (var i = 1; i <= 3; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i} contains secret information"));
        }

        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public void RedactByText_MultiplePages_ShouldRedactAllPages()
    {
        var pdfPath = CreateMultiPagePdf("test_redact_multi_page.pdf");
        var outputPath = CreateTestFilePath("test_redact_multi_page_output.pdf");
        var result = _tool.Execute(
            pdfPath,
            outputPath: outputPath,
            textToRedact: "secret");
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("3 occurrence", result);
        Assert.Contains("3 pages", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute(pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() => _tool.Execute(
            "nonexistent_file.pdf",
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void RedactByArea_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_redact_area.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            sessionId: sessionId,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50);
        Assert.Contains("Redaction applied", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.Single(document.Pages);
    }

    [Fact]
    public void RedactByText_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_redact_text.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            sessionId: sessionId,
            textToRedact: "redact");
        Assert.Contains("Redacted", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void RedactWithOverlay_WithSessionId_ShouldApplyOverlayInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_redact_overlay.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            sessionId: sessionId,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 50,
            overlayText: "[REDACTED]",
            fillColor: "Red");
        Assert.Contains("Redaction applied", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.Single(document.Pages);
    }

    #endregion
}