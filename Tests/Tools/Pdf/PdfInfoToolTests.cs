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

    private string CreateTestPdf(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} content"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var pdfPath = CreateTestPdf("test_get_content.pdf");
        var result = _tool.Execute("get_content", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("totalPages", out _));
        Assert.True(json.TryGetProperty("extractedPages", out _));
        Assert.True(json.TryGetProperty("content", out _));
        Assert.True(json.TryGetProperty("truncated", out _));
    }

    [Fact]
    public void GetContent_WithPageIndex_ShouldReturnSpecificPage()
    {
        var pdfPath = CreateTestPdf("test_get_content_page.pdf", 2);
        var result = _tool.Execute("get_content", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(1, json.GetProperty("pageIndex").GetInt32());
        Assert.Equal(2, json.GetProperty("totalPages").GetInt32());
        Assert.True(json.TryGetProperty("content", out _));
    }

    [SkippableFact]
    public void GetContent_WithMaxPages_ShouldLimitExtraction()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreateTestPdf("test_max_pages.pdf", 5);
        var result = _tool.Execute("get_content", pdfPath, maxPages: 2);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(2, json.GetProperty("extractedPages").GetInt32());
        Assert.True(json.GetProperty("truncated").GetBoolean());
        Assert.Equal(5, json.GetProperty("totalPages").GetInt32());
    }

    [Fact]
    public void GetContent_WithoutMaxPages_ShouldNotTruncate()
    {
        var pdfPath = CreateTestPdf("test_default_max.pdf");
        var result = _tool.Execute("get_content", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.False(json.GetProperty("truncated").GetBoolean());
    }

    [Fact]
    public void GetStatistics_ShouldReturnAllFields()
    {
        var pdfPath = CreateTestPdf("test_statistics.pdf");
        var result = _tool.Execute("get_statistics", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("fileSizeBytes", out _));
        Assert.True(json.TryGetProperty("fileSizeKb", out _));
        Assert.True(json.TryGetProperty("totalPages", out _));
        Assert.True(json.TryGetProperty("isEncrypted", out _));
        Assert.True(json.TryGetProperty("isLinearized", out _));
        Assert.True(json.TryGetProperty("bookmarks", out _));
        Assert.True(json.TryGetProperty("formFields", out _));
        Assert.True(json.TryGetProperty("totalAnnotations", out _));
        Assert.True(json.TryGetProperty("totalParagraphs", out _));
    }

    [Theory]
    [InlineData("GET_CONTENT")]
    [InlineData("Get_Content")]
    [InlineData("get_content")]
    public void Operation_ShouldBeCaseInsensitive_GetContent(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation.Replace("_", "")}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        Assert.Contains("content", result);
    }

    [Theory]
    [InlineData("GET_STATISTICS")]
    [InlineData("Get_Statistics")]
    [InlineData("get_statistics")]
    public void Operation_ShouldBeCaseInsensitive_GetStatistics(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_stat_{operation.Replace("_", "")}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        Assert.Contains("totalPages", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void GetContent_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_content", pdfPath, pageIndex: 99));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void GetStatistics_WithNoPath_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_statistics"));
        Assert.Contains("path", ex.Message.ToLower());
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_content"));
    }

    [Fact]
    public void GetContent_WithNegativeMaxPages_ShouldReturnNegativeExtracted()
    {
        var pdfPath = CreateTestPdf("test_negative_max.pdf");
        var result = _tool.Execute("get_content", pdfPath, maxPages: -1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("extractedPages").GetInt32() < 0);
    }

    #endregion

    #region Session

    [Fact]
    public void GetContent_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_content.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get_content", sessionId: sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("content", out _));
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_stats.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("totalPages", out _));
        Assert.True(json.TryGetProperty("note", out _));
    }

    [Fact]
    public void GetContent_WithSessionId_ShouldReflectInMemoryDocument()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(docBefore);

        var result = _tool.Execute("get_content", sessionId: sessionId);
        Assert.NotNull(result);

        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.Same(docBefore, docAfter);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_content", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_info.pdf");
        var pdfPath2 = CreateTestPdf("test_session_info.pdf", 3);
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get_content", pdfPath1, sessionId);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(3, json.GetProperty("totalPages").GetInt32());
    }

    #endregion
}