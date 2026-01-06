using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

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

    #region General

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
    public void Get_WithNoBookmarks_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No bookmarks found", result);
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

    [Fact]
    public void Edit_PageIndex_ShouldEditPageIndex()
    {
        var pdfPath = CreatePdfWithBookmark("test_edit_page.pdf");
        var outputPath = CreateTestFilePath("test_edit_page_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            bookmarkIndex: 1, pageIndex: 2);
        Assert.StartsWith("Edited bookmark", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            title: "Bookmark", pageIndex: 1);
        Assert.StartsWith("Added bookmark", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_get_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        Assert.Contains("count", result);
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
    public void Add_WithMissingTitle_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_title.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1));
        Assert.Contains("title is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, title: "Test"));
        Assert.Contains("pageIndex is required", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, title: "Test", pageIndex: 99));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingBookmarkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithBookmark("test_delete_no_index.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath));
        Assert.Contains("bookmarkIndex is required", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidBookmarkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, bookmarkIndex: 99));
        Assert.Contains("bookmarkIndex must be between", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingBookmarkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithBookmark("test_edit_no_index.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, title: "New Title"));
        Assert.Contains("bookmarkIndex is required", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidBookmarkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_edit_invalid_index.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, bookmarkIndex: 99, title: "Test"));
        Assert.Contains("bookmarkIndex must be between", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithBookmark("test_edit_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, bookmarkIndex: 1, pageIndex: 99));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

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
    public void Add_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute("add", sessionId: sessionId, title: "In-Memory Bookmark", pageIndex: 1);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.Outlines.Count > 0);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithBookmark("test_session_delete.pdf", "To Delete");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docBefore.Outlines.Count > 0);
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