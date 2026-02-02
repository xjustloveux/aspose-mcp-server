using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfRedactTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region Operation Routing

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute(pageIndex: 1, x: 100, y: 100, width: 200, height: 50));
    }

    #endregion

    #region File I/O Smoke Tests

    [Fact]
    public void RedactArea_ShouldRedactArea()
    {
        var pdfPath = CreateTestPdf("test_area.pdf");
        var outputPath = CreateTestFilePath("test_area_output.pdf");
        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        using var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count > 0);
    }

    [Fact]
    public void RedactArea_WithColorAndOverlay_ShouldApplyBoth()
    {
        var pdfPath = CreateTestPdf("test_both.pdf");
        var outputPath = CreateTestFilePath("test_both_output.pdf");
        var result = _tool.Execute(pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50,
            fillColor: "Red", overlayText: "CONFIDENTIAL");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        using var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count > 0);
    }

    [Fact]
    public void RedactByText_ShouldFindAndRedact()
    {
        const string textToFind = "redact";
        var pdfPath = CreateTestPdf("test_bytext.pdf", $"This text contains the word {textToFind} in it");
        var outputPath = CreateTestFilePath("test_bytext_output.pdf");

        var result = _tool.Execute(pdfPath, outputPath: outputPath, textToRedact: textToFind);

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        using var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count > 0);
    }

    [Fact]
    public void RedactByText_NotFound_ShouldReturnNoOccurrences()
    {
        var pdfPath = CreateTestPdf("test_notfound.pdf");
        var result = _tool.Execute(pdfPath, textToRedact: "nonexistent_12345");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("No occurrences of 'nonexistent_12345' found", data.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void RedactArea_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_area.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute(sessionId: sessionId,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 50);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.Pages.Count > 0);
    }

    [Fact]
    public void RedactByText_WithSessionId_ShouldModifyInMemory()
    {
        const string textToFind = "redact";
        var pdfPath = CreateTestPdf("test_session_text.pdf", $"Content to {textToFind} here");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute(sessionId: sessionId, textToRedact: textToFind);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
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
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
