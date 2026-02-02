using Aspose.Words;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Bookmark;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordBookmarkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordBookmarkToolTests : WordTestBase
{
    private readonly WordBookmarkTool _tool;

    public WordBookmarkToolTests()
    {
        _tool = new WordBookmarkTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddBookmark_ShouldAddBookmarkAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_add_bookmark.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_bookmark_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, name: "TestBookmark", text: "Bookmarked text");
        var doc = new Document(outputPath);
        var bookmark = doc.Range.Bookmarks["TestBookmark"];
        Assert.NotNull(bookmark);
    }

    [Fact]
    public void GetBookmarks_ShouldReturnBookmarksFromFile()
    {
        var docPath = CreateWordDocument("test_get_bookmarks.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("Bookmark1");
        builder.Write("Content1");
        builder.EndBookmark("Bookmark1");
        doc.Save(docPath);

        var result = _tool.Execute("get", docPath);
        var data = GetResultData<GetBookmarksResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Bookmarks, b => b.Name == "Bookmark1");
    }

    [Fact]
    public void DeleteBookmark_ShouldDeleteBookmarkAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_bookmark.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("BookmarkToDelete");
        builder.Write("Content");
        builder.EndBookmark("BookmarkToDelete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_bookmark_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, name: "BookmarkToDelete", keepText: true);
        var resultDoc = new Document(outputPath);
        Assert.Null(resultDoc.Range.Bookmarks["BookmarkToDelete"]);
    }

    [Fact]
    public void EditBookmark_ShouldEditBookmarkAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_edit_bookmark.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("BookmarkToEdit");
        builder.Write("Original text");
        builder.EndBookmark("BookmarkToEdit");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_bookmark_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, name: "BookmarkToEdit", newText: "Updated text");
        var resultDoc = new Document(outputPath);
        var bookmark = resultDoc.Range.Bookmarks["BookmarkToEdit"];
        Assert.NotNull(bookmark);
        Assert.Contains("Updated text", bookmark.Text);
    }

    [Fact]
    public void GotoBookmark_ShouldNavigateToBookmark()
    {
        var docPath = CreateWordDocumentWithContent("test_goto_bookmark.docx", "Content before bookmark");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TargetBookmark");
        builder.Write("Target content");
        builder.EndBookmark("TargetBookmark");
        doc.Save(docPath);

        var result = _tool.Execute("goto", docPath, name: "TargetBookmark");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("TargetBookmark", data.Message);
        Assert.Contains("Target content", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, name: $"BM_{operation}", text: "Text");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var doc = new Document(outputPath);
        Assert.NotNull(doc.Range.Bookmarks[$"BM_{operation}"]);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddBookmark_WithSessionId_ShouldAddBookmarkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add.docx", "Test content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, name: "SessionBookmark", text: "Session text");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var bookmark = doc.Range.Bookmarks["SessionBookmark"];
        Assert.NotNull(bookmark);
    }

    [Fact]
    public void GetBookmarks_WithSessionId_ShouldReturnBookmarks()
    {
        var docPath = CreateWordDocument("test_session_get.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("ExistingBookmark");
        builder.EndBookmark("ExistingBookmark");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetBookmarksResult>(result);
        Assert.Contains(data.Bookmarks, b => b.Name == "ExistingBookmark");
        var output = GetResultOutput<GetBookmarksResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void EditBookmark_WithSessionId_ShouldEditInMemory()
    {
        var docPath = CreateWordDocument("test_session_edit.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("EditMe");
        builder.Write("Original");
        builder.EndBookmark("EditMe");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("edit", sessionId: sessionId, name: "EditMe", newText: "Updated via session");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var bookmark = sessionDoc.Range.Bookmarks["EditMe"];
        Assert.Contains("Updated via session", bookmark.Text);
    }

    [Fact]
    public void DeleteBookmark_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("DeleteMe");
        builder.EndBookmark("DeleteMe");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, name: "DeleteMe");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Null(sessionDoc.Range.Bookmarks["DeleteMe"]);
    }

    [Fact]
    public void GotoBookmark_WithSessionId_ShouldNavigate()
    {
        var docPath = CreateWordDocument("test_session_goto.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("NavigateHere");
        builder.Write("Target content");
        builder.EndBookmark("NavigateHere");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("goto", sessionId: sessionId, name: "NavigateHere");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("NavigateHere", data.Message);
        Assert.Contains("Target content", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path.docx");
        var doc1 = new Document(docPath1);
        var builder1 = new DocumentBuilder(doc1);
        builder1.StartBookmark("PathBookmark");
        builder1.EndBookmark("PathBookmark");
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session.docx");
        var doc2 = new Document(docPath2);
        var builder2 = new DocumentBuilder(doc2);
        builder2.StartBookmark("SessionBookmark");
        builder2.EndBookmark("SessionBookmark");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);
        var data = GetResultData<GetBookmarksResult>(result);
        Assert.Contains(data.Bookmarks, b => b.Name == "SessionBookmark");
        Assert.DoesNotContain(data.Bookmarks, b => b.Name == "PathBookmark");
    }

    #endregion
}
