using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordBookmarkToolTests : WordTestBase
{
    private readonly WordBookmarkTool _tool;

    public WordBookmarkToolTests()
    {
        _tool = new WordBookmarkTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddBookmark_ShouldAddBookmark()
    {
        var docPath = CreateWordDocumentWithContent("test_add_bookmark.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_bookmark_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, name: "TestBookmark", text: "Bookmarked text");
        var doc = new Document(outputPath);
        var bookmark = doc.Range.Bookmarks["TestBookmark"];
        Assert.NotNull(bookmark);
        Assert.Contains("Bookmarked text", bookmark.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetBookmarks_ShouldReturnAllBookmarks()
    {
        var docPath = CreateWordDocument("test_get_bookmarks.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("Bookmark1");
        builder.EndBookmark("Bookmark1");
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Bookmark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteBookmark_ShouldDeleteBookmark()
    {
        var docPath = CreateWordDocument("test_delete_bookmark.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("BookmarkToDelete");
        builder.EndBookmark("BookmarkToDelete");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_bookmark_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, name: "BookmarkToDelete", keepText: true);
        var resultDoc = new Document(outputPath);
        // Verify bookmark was deleted - try to access it, should throw or return null
        Bookmark? bookmark = null;
        try
        {
            bookmark = resultDoc.Range.Bookmarks["BookmarkToDelete"];
        }
        catch
        {
            // Bookmark not found, which is expected
        }

        Assert.Null(bookmark);
    }

    [Fact]
    public void EditBookmark_ShouldEditBookmark()
    {
        var docPath = CreateWordDocument("test_edit_bookmark.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("BookmarkToEdit");
        builder.Write("Original text");
        builder.EndBookmark("BookmarkToEdit");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_bookmark_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, name: "BookmarkToEdit", text: "Updated text");
        var resultDoc = new Document(outputPath);
        var bookmark = resultDoc.Range.Bookmarks["BookmarkToEdit"];
        Assert.NotNull(bookmark);
        Assert.Contains("Updated text", bookmark.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GotoBookmark_ShouldNavigateToBookmark()
    {
        var docPath = CreateWordDocumentWithContent("test_goto_bookmark.docx", "Content before bookmark");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Content after bookmark");
        builder.StartBookmark("TargetBookmark");
        builder.EndBookmark("TargetBookmark");
        doc.Save(docPath);
        var result = _tool.Execute("goto", docPath, name: "TargetBookmark");
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("TargetBookmark", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void AddBookmark_WithEmptyName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_empty_name.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, name: "", text: "Some text"));

        Assert.Contains("Bookmark name is required", ex.Message);
    }

    [Fact]
    public void AddBookmark_WithNullName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_null_name.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_null_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, name: null, text: "Some text"));

        Assert.Contains("Bookmark name is required", ex.Message);
    }

    [Fact]
    public void AddBookmark_WithDuplicateName_ShouldThrowInvalidOperationException()
    {
        var docPath = CreateWordDocument("test_add_duplicate.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("ExistingBookmark");
        builder.Write("Existing content");
        builder.EndBookmark("ExistingBookmark");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_duplicate_output.docx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, name: "ExistingBookmark", text: "New text"));

        Assert.Contains("already exists", ex.Message);
        Assert.Contains("ExistingBookmark", ex.Message);
    }

    [Fact]
    public void AddBookmark_WithDuplicateNameCaseInsensitive_ShouldThrowInvalidOperationException()
    {
        var docPath = CreateWordDocument("test_add_duplicate_case.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("MyBookmark");
        builder.EndBookmark("MyBookmark");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_duplicate_case_output.docx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, name: "MYBOOKMARK", text: "Text"));

        Assert.Contains("already exists", ex.Message);
        Assert.Contains("case-insensitive", ex.Message);
    }

    [Fact]
    public void AddBookmark_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_invalid_para.docx", "Single paragraph");
        var outputPath = CreateTestFilePath("test_add_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, name: "NewBookmark",
                text: "Text", paragraphIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditBookmark_WithEmptyName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_empty_name.docx", "Test content");
        var outputPath = CreateTestFilePath("test_edit_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, name: "", text: "New text"));

        Assert.Contains("Bookmark name is required", ex.Message);
    }

    [Fact]
    public void EditBookmark_WithNonExistentBookmark_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_nonexistent.docx", "Test content");
        var outputPath = CreateTestFilePath("test_edit_nonexistent_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, name: "NonExistentBookmark", text: "New text"));

        Assert.Contains("not found", ex.Message);
        Assert.Contains("NonExistentBookmark", ex.Message);
    }

    [Fact]
    public void EditBookmark_WithNoChanges_ShouldReturnNoChangesMessage()
    {
        var docPath = CreateWordDocument("test_edit_no_changes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TestBookmark");
        builder.Write("Original text");
        builder.EndBookmark("TestBookmark");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_no_changes_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, name: "TestBookmark"));

        Assert.Contains("newName or newText is required", ex.Message);
    }

    [Fact]
    public void EditBookmark_WithDuplicateNewName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_dup_name.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("Bookmark1");
        builder.EndBookmark("Bookmark1");
        builder.StartBookmark("Bookmark2");
        builder.EndBookmark("Bookmark2");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_dup_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, name: "Bookmark1", newName: "Bookmark2"));

        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void DeleteBookmark_WithEmptyName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_empty_name.docx", "Test content");
        var outputPath = CreateTestFilePath("test_delete_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, name: ""));

        Assert.Contains("Bookmark name is required", ex.Message);
    }

    [Fact]
    public void DeleteBookmark_WithNonExistentBookmark_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_nonexistent.docx", "Test content");
        var outputPath = CreateTestFilePath("test_delete_nonexistent_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, name: "NonExistentBookmark"));

        Assert.Contains("not found", ex.Message);
        Assert.Contains("NonExistentBookmark", ex.Message);
    }

    [Fact]
    public void GotoBookmark_WithEmptyName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_goto_empty_name.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("goto", docPath, name: ""));

        Assert.Contains("Bookmark name is required", ex.Message);
    }

    [Fact]
    public void GotoBookmark_WithNonExistentBookmark_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_goto_nonexistent.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("goto", docPath, name: "NonExistentBookmark"));

        Assert.Contains("not found", ex.Message);
        Assert.Contains("NonExistentBookmark", ex.Message);
    }

    [Fact]
    public void GetBookmarks_WithNoBookmarks_ShouldReturnEmptyResult()
    {
        var docPath = CreateWordDocumentWithContent("test_get_no_bookmarks.docx", "No bookmarks here");
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.Contains("\"count\":0", result.Replace(" ", ""));
        Assert.Contains("No bookmarks found", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddBookmark_WithSessionId_ShouldAddBookmarkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add.docx", "Test content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, name: "SessionBookmark", text: "Session text");
        Assert.Contains("SessionBookmark", result);

        // Verify in-memory document has the bookmark
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
        Assert.Contains("ExistingBookmark", result);
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

        // Assert - verify in-memory change
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

        // Assert - verify in-memory deletion
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var bookmark = sessionDoc.Range.Bookmarks["DeleteMe"];
        Assert.Null(bookmark);
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
        Assert.Contains("NavigateHere", result);
        Assert.Contains("Target content", result);
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("get", docPath1, sessionId);

        // Assert - should use sessionId, returning SessionBookmark not PathBookmark
        Assert.Contains("SessionBookmark", result);
        Assert.DoesNotContain("PathBookmark", result);
    }

    #endregion
}