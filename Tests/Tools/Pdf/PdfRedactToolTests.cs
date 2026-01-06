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

    private string CreateTestPdf(string fileName, string content = "Text to redact")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment(content));
        document.Save(filePath);
        return filePath;
    }

    private string CreateMultiPagePdf(string fileName, int pageCount = 3)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 1; i <= pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i} contains secret information"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void RedactArea_ShouldRedactArea()
    {
        var pdfPath = CreateTestPdf("test_area.pdf");
        var outputPath = CreateTestFilePath("test_area_output.pdf");
        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redaction applied", result);
    }

    [Fact]
    public void RedactArea_WithColor_ShouldApplyColor()
    {
        var pdfPath = CreateTestPdf("test_color.pdf");
        var outputPath = CreateTestFilePath("test_color_output.pdf");
        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50, fillColor: "255,0,0");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redaction applied", result);
    }

    [Fact]
    public void RedactArea_WithColorName_ShouldApplyColor()
    {
        var pdfPath = CreateTestPdf("test_colorname.pdf");
        var outputPath = CreateTestFilePath("test_colorname_output.pdf");
        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50, fillColor: "Blue");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redaction applied", result);
    }

    [Fact]
    public void RedactArea_WithOverlayText_ShouldApplyOverlay()
    {
        var pdfPath = CreateTestPdf("test_overlay.pdf");
        var outputPath = CreateTestFilePath("test_overlay_output.pdf");
        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50, overlayText: "[REDACTED]");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redaction applied", result);
    }

    [Fact]
    public void RedactArea_WithColorAndOverlay_ShouldApplyBoth()
    {
        var pdfPath = CreateTestPdf("test_both.pdf");
        var outputPath = CreateTestFilePath("test_both_output.pdf");
        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50,
            fillColor: "Red", overlayText: "CONFIDENTIAL");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redaction applied", result);
    }

    [Fact]
    public void RedactArea_WithoutOutputPath_ShouldOverwriteInput()
    {
        var pdfPath = CreateTestPdf("test_overwrite.pdf");
        var result = _tool.Execute(pdfPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50);
        Assert.True(File.Exists(pdfPath));
        Assert.StartsWith("Redaction applied", result);
    }

    [Fact]
    public void RedactByText_ShouldFindAndRedact()
    {
        const string textToFind = "redact";
        var pdfPath = CreateTestPdf("test_bytext.pdf", $"This text contains the word {textToFind} in it");
        var outputPath = CreateTestFilePath("test_bytext_output.pdf");

        var result = _tool.Execute(pdfPath, outputPath: outputPath, textToRedact: textToFind);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redacted 1 occurrence", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        var extractedText = textAbsorber.Text;

        Assert.DoesNotContain(textToFind, extractedText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RedactByText_OnSpecificPage_ShouldRedactOnlyThatPage()
    {
        const string textToFind = "Text";
        var pdfPath = CreateTestPdf("test_bytext_page.pdf", $"{textToFind} to redact on page");
        var outputPath = CreateTestFilePath("test_bytext_page_output.pdf");

        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, textToRedact: textToFind);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redacted 1 occurrence", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.DoesNotContain(textToFind, textAbsorber.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RedactByText_CaseInsensitive_ShouldFindText()
    {
        var pdfPath = CreateTestPdf("test_case.pdf", "This contains text to find");
        var outputPath = CreateTestFilePath("test_case_output.pdf");

        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            textToRedact: "TEXT", caseSensitive: false);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redacted 1 occurrence", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.DoesNotContain("text", textAbsorber.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RedactByText_NotFound_ShouldReturnNoOccurrences()
    {
        var pdfPath = CreateTestPdf("test_notfound.pdf");

        var result = _tool.Execute(pdfPath, textToRedact: "nonexistent_12345");

        Assert.Equal("No occurrences of 'nonexistent_12345' found. No redactions applied.", result);
    }

    [Fact]
    public void RedactByText_WithOverlay_ShouldApplyOverlay()
    {
        const string textToFind = "Text";
        const string overlayText = "[CLASSIFIED]";
        var pdfPath = CreateTestPdf("test_text_overlay.pdf", $"{textToFind} to be replaced");
        var outputPath = CreateTestFilePath("test_text_overlay_output.pdf");

        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            textToRedact: textToFind, overlayText: overlayText);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redacted 1 occurrence", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.DoesNotContain(textToFind, textAbsorber.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RedactByText_WithColor_ShouldApplyColor()
    {
        const string textToFind = "redact";
        var pdfPath = CreateTestPdf("test_text_color.pdf", $"Content to {textToFind} here");
        var outputPath = CreateTestFilePath("test_text_color_output.pdf");

        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            textToRedact: textToFind, fillColor: "Red");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redacted 1 occurrence", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.DoesNotContain(textToFind, textAbsorber.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RedactByText_MultiplePages_ShouldRedactAllPages()
    {
        var pdfPath = CreateMultiPagePdf("test_multipage.pdf");
        var outputPath = CreateTestFilePath("test_multipage_output.pdf");

        var result = _tool.Execute(pdfPath, outputPath: outputPath, textToRedact: "secret");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Redacted 3 occurrence(s) of 'secret' on 3 pages", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.DoesNotContain("secret", textAbsorber.Text, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception

    [Fact]
    public void RedactArea_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pdfPath, pageIndex: 99, x: 100, y: 100, width: 200, height: 50));
        Assert.StartsWith("pageIndex must be between 1 and", ex.Message);
    }

    [Fact]
    public void RedactArea_WithMissingPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_no_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pdfPath, x: 100, y: 100, width: 200, height: 50));
        Assert.Equal("pageIndex is required for area redaction", ex.Message);
    }

    [Fact]
    public void RedactArea_WithMissingX_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_no_x.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pdfPath, pageIndex: 1, y: 100, width: 200, height: 50));
        Assert.Equal("x is required for area redaction", ex.Message);
    }

    [Fact]
    public void RedactArea_WithMissingY_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_no_y.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pdfPath, pageIndex: 1, x: 100, width: 200, height: 50));
        Assert.Equal("y is required for area redaction", ex.Message);
    }

    [Fact]
    public void RedactArea_WithMissingWidth_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_no_width.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pdfPath, pageIndex: 1, x: 100, y: 100, height: 50));
        Assert.Equal("width is required for area redaction", ex.Message);
    }

    [Fact]
    public void RedactArea_WithMissingHeight_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_no_height.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pdfPath, pageIndex: 1, x: 100, y: 100, width: 200));
        Assert.Equal("height is required for area redaction", ex.Message);
    }

    [Fact]
    public void RedactByText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_text_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pdfPath, pageIndex: 99, textToRedact: "Text"));
        Assert.StartsWith("pageIndex must be between 1 and", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute(pageIndex: 1, x: 100, y: 100, width: 200, height: 50));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("nonexistent_file.pdf", pageIndex: 1, x: 100, y: 100, width: 200, height: 50));
    }

    #endregion

    #region Session

    [Fact]
    public void RedactArea_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_area.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute(sessionId: sessionId,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50);

        Assert.StartsWith("Redaction applied", result);
        Assert.Contains(sessionId, result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void RedactArea_WithSessionId_AndOptions_ShouldApplyOptionsInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_options.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute(sessionId: sessionId,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50,
            overlayText: "[REDACTED]", fillColor: "Red");

        Assert.StartsWith("Redaction applied", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void RedactByText_WithSessionId_ShouldModifyInMemory()
    {
        const string textToFind = "redact";
        var pdfPath = CreateTestPdf("test_session_text.pdf", $"Content to {textToFind} here");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute(sessionId: sessionId, textToRedact: textToFind);

        Assert.StartsWith("Redacted 1 occurrence", result);
        Assert.Contains(sessionId, result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);

        var textAbsorber = new TextAbsorber();
        document.Pages.Accept(textAbsorber);
        Assert.DoesNotContain(textToFind, textAbsorber.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute(sessionId: "invalid_session", pageIndex: 1, x: 100, y: 100, width: 200, height: 50));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_redact.pdf", "Path content");
        var pdfPath2 = CreateTestPdf("test_session_redact.pdf", "Session content");
        var sessionId = OpenSession(pdfPath2);

        var result = _tool.Execute(pdfPath1, sessionId,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50);

        Assert.Contains(sessionId, result);
    }

    #endregion
}