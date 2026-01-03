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

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddBookmark_ShouldAddBookmark()
    {
        var pdfPath = CreateTestPdf("test_add_bookmark.pdf");
        var outputPath = CreateTestFilePath("test_add_bookmark_output.pdf");
        var result = _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            title: "Chapter 1",
            pageIndex: 1);
        Assert.Contains("Added bookmark", result);
        using var document = new Document(outputPath);
        Assert.True(document.Outlines.Count > 0, "Document should contain at least one bookmark");
        Assert.Equal("Chapter 1", document.Outlines[1].Title);
    }

    [Fact]
    public void AddBookmark_InvalidPageIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            title: "Test",
            pageIndex: 99));
    }

    [Fact]
    public void GetBookmarks_WithBookmarks_ShouldReturnBookmarkInfo()
    {
        var pdfPath = CreateTestPdf("test_get_bookmarks.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Test Bookmark",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 1", result);
        Assert.Contains("Test Bookmark", result);
    }

    [Fact]
    public void GetBookmarks_Empty_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No bookmarks found", result);
    }

    [Fact]
    public void DeleteBookmark_ShouldDeleteBookmark()
    {
        var pdfPath = CreateTestPdf("test_delete_bookmark.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Bookmark to Delete",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_delete_bookmark_output.pdf");
        var result = _tool.Execute(
            "delete",
            pdfPath,
            outputPath: outputPath,
            bookmarkIndex: 1);
        Assert.Contains("Deleted bookmark", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.Outlines);
    }

    [Fact]
    public void DeleteBookmark_InvalidIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_delete_invalid.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            bookmarkIndex: 99));
    }

    [Fact]
    public void EditBookmark_ShouldEditTitle()
    {
        var pdfPath = CreateTestPdf("test_edit_bookmark.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Original Title",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_bookmark_output.pdf");
        var result = _tool.Execute(
            "edit",
            pdfPath,
            outputPath: outputPath,
            bookmarkIndex: 1,
            title: "Updated Title");
        Assert.Contains("Edited bookmark", result);
        using var resultDocument = new Document(outputPath);
        Assert.Equal("Updated Title", resultDocument.Outlines[1].Title);
    }

    [Fact]
    public void EditBookmark_ShouldEditPageIndex()
    {
        var pdfPath = CreateTestPdf("test_edit_page.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Bookmark",
                Action = new GoToAction(document.Pages[1])
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_page_output.pdf");
        var result = _tool.Execute(
            "edit",
            pdfPath,
            outputPath: outputPath,
            bookmarkIndex: 1,
            pageIndex: 2);
        Assert.Contains("Edited bookmark", result);
    }

    [Fact]
    public void EditBookmark_InvalidBookmarkIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_edit_invalid.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            pdfPath,
            bookmarkIndex: 99,
            title: "Test"));
    }

    [Fact]
    public void EditBookmark_InvalidPageIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_edit_invalid_page.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Bookmark",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            pdfPath,
            bookmarkIndex: 1,
            pageIndex: 99));
    }

    [Fact]
    public void UnknownOperation_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_exception_unknown.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("invalid_operation", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void AddBookmark_WithMissingTitle_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_missing_title.pdf");

        // Act & Assert - missing title
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetBookmarks_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        using (var document = new Document(pdfPath))
        {
            var outlineItem = new OutlineItemCollection(document.Outlines)
            {
                Title = "Session Bookmark",
                Destination = new XYZExplicitDestination(document.Pages[1], 0, 0, 1)
            };
            document.Outlines.Add(outlineItem);
            document.Save(pdfPath);
        }

        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("Session Bookmark", result);
    }

    [Fact]
    public void AddBookmark_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            title: "New Session Bookmark",
            pageIndex: 1);
        Assert.Contains("Added bookmark", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void AddBookmark_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute(
            "add",
            sessionId: sessionId,
            title: "In-Memory Bookmark",
            pageIndex: 1);

        // Assert - verify in-memory changes
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.Outlines.Count > 0);
    }

    #endregion
}