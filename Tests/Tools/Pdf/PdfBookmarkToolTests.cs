using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Pdf.Bookmark;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfBookmarkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfBookmarkToolTests : PdfTestBase
{
    private readonly PdfBookmarkTool _tool;

    public PdfBookmarkToolTests()
    {
        _tool = new PdfBookmarkTool(SessionManager);
    }

    private string CreateTestPdf(string fileName, int pageCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithBookmark(string fileName, string bookmarkTitle = "Test Bookmark")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var outlineItem = new OutlineItemCollection(document.Outlines)
        {
            Title = bookmarkTitle,
            Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
        };
        document.Outlines.Add(outlineItem);
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddBookmark()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            title: "Chapter 1", pageIndex: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Added bookmark", data.Message);
        using var document = new Document(outputPath);
        Assert.True(document.Outlines.Count > 0);
        Assert.Equal("Chapter 1", document.Outlines[1].Title);
    }

    [Fact]
    public void Get_WithBookmarks_ShouldReturnBookmarkInfo()
    {
        var pdfPath = CreatePdfWithBookmark("test_get.pdf");
        var result = _tool.Execute("get", pdfPath);
        var data = GetResultData<GetBookmarksPdfResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Contains(data.Items, b => b.Title == "Test Bookmark");
    }

    [Fact]
    public void Delete_ShouldDeleteBookmark()
    {
        var pdfPath = CreatePdfWithBookmark("test_delete.pdf", "Bookmark to Delete");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath, bookmarkIndex: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Deleted bookmark", data.Message);
        using var document = new Document(outputPath);
        Assert.Empty(document.Outlines);
    }

    [Fact]
    public void Edit_Title_ShouldEditTitle()
    {
        var pdfPath = CreatePdfWithBookmark("test_edit_title.pdf", "Original Title");
        var outputPath = CreateTestFilePath("test_edit_title_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            bookmarkIndex: 1, title: "Updated Title");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Edited bookmark", data.Message);
        using var document = new Document(outputPath);
        Assert.Equal("Updated Title", document.Outlines[1].Title);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            title: "Bookmark", pageIndex: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Added bookmark", data.Message);
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
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreatePdfWithBookmark("test_session_get.pdf", "Session Bookmark");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetBookmarksPdfResult>(result);
        Assert.Contains(data.Items, b => b.Title == "Session Bookmark");
        var output = GetResultOutput<GetBookmarksPdfResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            title: "Session Bookmark", pageIndex: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Added bookmark", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithBookmark("test_session_delete.pdf", "To Delete");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("delete", sessionId: sessionId, bookmarkIndex: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Deleted bookmark", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.Empty(docAfter.Outlines);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInSession()
    {
        var pdfPath = CreatePdfWithBookmark("test_session_edit.pdf", "Original");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("edit", sessionId: sessionId,
            bookmarkIndex: 1, title: "Edited");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Edited bookmark", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("Edited", document.Outlines[1].Title);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_file.pdf");
        var pdfPath2 = CreatePdfWithBookmark("test_session_file.pdf", "Session Bookmark");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId);
        var data = GetResultData<GetBookmarksPdfResult>(result);
        Assert.Contains(data.Items, b => b.Title == "Session Bookmark");
    }

    #endregion
}
