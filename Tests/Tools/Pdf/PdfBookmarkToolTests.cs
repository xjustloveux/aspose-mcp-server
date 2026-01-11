using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
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
        Assert.StartsWith("Added bookmark", result);
        using var document = new Document(outputPath);
        Assert.True(document.Outlines.Count > 0);
        Assert.Equal("Chapter 1", document.Outlines[1].Title);
    }

    [Fact]
    public void Get_WithBookmarks_ShouldReturnBookmarkInfo()
    {
        var pdfPath = CreatePdfWithBookmark("test_get.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 1", result);
        Assert.Contains("Test Bookmark", result);
    }

    [Fact]
    public void Delete_ShouldDeleteBookmark()
    {
        var pdfPath = CreatePdfWithBookmark("test_delete.pdf", "Bookmark to Delete");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath, bookmarkIndex: 1);
        Assert.StartsWith("Deleted bookmark", result);
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
        Assert.StartsWith("Edited bookmark", result);
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
        Assert.StartsWith("Added bookmark", result);
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
        Assert.Contains("Session Bookmark", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            title: "Session Bookmark", pageIndex: 1);
        Assert.StartsWith("Added bookmark", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithBookmark("test_session_delete.pdf", "To Delete");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("delete", sessionId: sessionId, bookmarkIndex: 1);
        Assert.StartsWith("Deleted bookmark", result);
        Assert.Contains("session", result);
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
        Assert.StartsWith("Edited bookmark", result);
        Assert.Contains("session", result);
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
        Assert.Contains("Session Bookmark", result);
    }

    #endregion
}
