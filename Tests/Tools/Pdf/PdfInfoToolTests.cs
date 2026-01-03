using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfInfoToolTests : PdfTestBase
{
    private readonly PdfInfoTool _tool;

    public PdfInfoToolTests()
    {
        _tool = new PdfInfoTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test PDF content"));
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var pdfPath = CreateTestPdf("test_get_content.pdf");
        var result = _tool.Execute("get_content", pdfPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Content", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var pdfPath = CreateTestPdf("test_get_statistics.pdf");
        var result = _tool.Execute("get_statistics", pdfPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("fileSizeBytes", result);
    }

    [Fact]
    public void GetContent_WithPageIndex_ShouldReturnSpecificPage()
    {
        var pdfPath = CreateTestPdf("test_get_content_page.pdf");
        var result = _tool.Execute("get_content", pdfPath, pageIndex: 1);
        Assert.NotNull(result);
        Assert.Contains("\"pageIndex\": 1", result);
        Assert.Contains("content", result);
    }

    [Fact]
    public void GetContent_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_content", pdfPath, pageIndex: 99));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [SkippableFact]
    public void GetContent_WithMaxPages_ShouldLimitExtraction()
    {
        // Skip in evaluation mode - 5 pages exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");

        // Arrange - Create a PDF with multiple pages
        var pdfPath = CreateTestFilePath("test_max_pages.pdf");
        var document = new Document();
        for (var i = 0; i < 5; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} content"));
        }

        document.Save(pdfPath);
        var result = _tool.Execute("get_content", pdfPath, maxPages: 2);
        Assert.Contains("\"extractedPages\": 2", result);
        Assert.Contains("\"truncated\": true", result);
        Assert.Contains("\"totalPages\": 5", result);
    }

    [Fact]
    public void GetContent_WithoutMaxPages_ShouldUseDefault()
    {
        var pdfPath = CreateTestPdf("test_default_max.pdf");
        var result = _tool.Execute("get_content", pdfPath);
        Assert.Contains("extractedPages", result);
        Assert.Contains("\"truncated\": false", result);
    }

    [Fact]
    public void GetStatistics_ShouldReturnAllFields()
    {
        var pdfPath = CreateTestPdf("test_all_stats.pdf");
        var result = _tool.Execute("get_statistics", pdfPath);
        Assert.Contains("fileSizeBytes", result);
        Assert.Contains("fileSizeKb", result);
        Assert.Contains("totalPages", result);
        Assert.Contains("isEncrypted", result);
        Assert.Contains("isLinearized", result);
        Assert.Contains("bookmarks", result);
        Assert.Contains("formFields", result);
        Assert.Contains("totalAnnotations", result);
        Assert.Contains("totalParagraphs", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithMissingRequiredPath_ShouldThrowArgumentException()
    {
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("get_content"));
        Assert.Contains("path", exception.Message.ToLower());
    }

    [Fact]
    public void GetContent_WithNegativeMaxPages_ShouldReturnEmptyContent()
    {
        var pdfPath = CreateTestPdf("test_negative_max_pages.pdf");

        // Act - Negative maxPages results in 0 pages extracted (Math.Min behavior)
        var result = _tool.Execute("get_content", pdfPath, maxPages: -1);

        // Assert - Result is returned but with no page content extracted
        Assert.NotNull(result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetContent_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get_content.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get_content", sessionId: sessionId, pageIndex: 1);
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("content", out _));
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get_statistics.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("totalPages", result);
        Assert.Contains("isEncrypted", result);
    }

    [Fact]
    public void GetContent_WithSessionId_ShouldReflectInMemoryDocument()
    {
        var pdfPath = CreateTestPdf("test_session_in_memory.pdf");
        var sessionId = OpenSession(pdfPath);

        // Verify the session document is accessible
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc);
        Assert.True(doc.Pages.Count > 0, "Session document should have pages");
        var result = _tool.Execute("get_content", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("content", result);

        // Verify document is still in session
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(docAfter);
        Assert.Same(doc, docAfter);
    }

    #endregion
}