using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results.Pdf.Info;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfInfoTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var pdfPath = CreateTestPdf("test_get_content.pdf");
        var result = _tool.Execute("get_content", pdfPath);
        var data = GetResultData<GetPdfContentResult>(result);
        Assert.True(data.TotalPages > 0);
        Assert.NotNull(data.ExtractedPages);
        Assert.NotNull(data.Content);
    }

    [Fact]
    public void GetContent_WithPageIndex_ShouldReturnSpecificPage()
    {
        var pdfPath = CreateTestPdf("test_get_content_page.pdf", 2);
        var result = _tool.Execute("get_content", pdfPath, pageIndex: 1);
        var data = GetResultData<GetPdfContentResult>(result);
        Assert.Equal(1, data.PageIndex);
        Assert.Equal(2, data.TotalPages);
    }

    [Fact]
    public void GetStatistics_ShouldReturnAllFields()
    {
        var pdfPath = CreateTestPdf("test_statistics.pdf");
        var result = _tool.Execute("get_statistics", pdfPath);
        var data = GetResultData<GetPdfStatisticsResult>(result);
        Assert.NotNull(data.FileSizeBytes);
        Assert.True(data.TotalPages > 0);
        Assert.False(data.IsEncrypted);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_CONTENT")]
    [InlineData("Get_Content")]
    [InlineData("get_content")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation.Replace("_", "")}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        var data = GetResultData<GetPdfContentResult>(result);
        Assert.NotNull(data.Content);
        Assert.True(data.TotalPages > 0);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetContent_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_content.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get_content", sessionId: sessionId, pageIndex: 1);
        var data = GetResultData<GetPdfContentResult>(result);
        Assert.NotNull(data.Content);
        var output = GetResultOutput<GetPdfContentResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_stats.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        var data = GetResultData<GetPdfStatisticsResult>(result);
        Assert.True(data.TotalPages > 0);
        var output = GetResultOutput<GetPdfStatisticsResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<GetPdfContentResult>(result);
        Assert.Equal(3, data.TotalPages);
        var output = GetResultOutput<GetPdfContentResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
