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

    #region General

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
    public void AddBookmark_WithParagraphIndex_ShouldAddAtSpecificParagraph()
    {
        var docPath = CreateWordDocumentWithContent("test_add_para_idx.docx", "First paragraph");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Second paragraph");
        builder.Writeln("Third paragraph");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_para_idx_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, name: "ParaBookmark", text: "At para 1",
            paragraphIndex: 1);
        var resultDoc = new Document(outputPath);
        var bookmark = resultDoc.Range.Bookmarks["ParaBookmark"];
        Assert.NotNull(bookmark);
    }

    [Fact]
    public void AddBookmark_WithNegativeOneParagraphIndex_ShouldAddAtBeginning()
    {
        var docPath = CreateWordDocumentWithContent("test_add_begin.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_begin_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath, name: "BeginBookmark", text: "At beginning",
            paragraphIndex: -1);
        Assert.StartsWith("Bookmark added successfully", result);
        var doc = new Document(outputPath);
        Assert.NotNull(doc.Range.Bookmarks["BeginBookmark"]);
    }

    [Fact]
    public void AddBookmark_WithoutParagraphIndex_ShouldAddAtEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_add_end.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_end_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath, name: "EndBookmark", text: "At end");
        Assert.StartsWith("Bookmark added successfully", result);
    }

    [Fact]
    public void AddBookmark_WithoutText_ShouldAddEmptyBookmark()
    {
        var docPath = CreateWordDocumentWithContent("test_add_empty.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_empty_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, name: "EmptyBookmark");
        var doc = new Document(outputPath);
        var bookmark = doc.Range.Bookmarks["EmptyBookmark"];
        Assert.NotNull(bookmark);
        Assert.Empty(bookmark.Text);
    }

    [Fact]
    public void GetBookmarks_ShouldReturnAllBookmarks()
    {
        var docPath = CreateWordDocument("test_get_bookmarks.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("Bookmark1");
        builder.Write("Content1");
        builder.EndBookmark("Bookmark1");
        builder.StartBookmark("Bookmark2");
        builder.Write("Content2");
        builder.EndBookmark("Bookmark2");
        doc.Save(docPath);

        var result = _tool.Execute("get", docPath);
        Assert.Contains("Bookmark1", result);
        Assert.Contains("Bookmark2", result);
        Assert.Contains("\"count\":2", result.Replace(" ", ""));
    }

    [Fact]
    public void GetBookmarks_ShouldReturnJsonWithCorrectFields()
    {
        var docPath = CreateWordDocument("test_get_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TestBookmark");
        builder.Write("Test content");
        builder.EndBookmark("TestBookmark");
        doc.Save(docPath);

        var result = _tool.Execute("get", docPath);
        Assert.Contains("\"name\"", result);
        Assert.Contains("\"text\"", result);
        Assert.Contains("\"length\"", result);
        Assert.Contains("\"index\"", result);
    }

    [Fact]
    public void GetBookmarks_WithNoBookmarks_ShouldReturnEmptyResult()
    {
        var docPath = CreateWordDocumentWithContent("test_get_no_bookmarks.docx", "No bookmarks here");
        var result = _tool.Execute("get", docPath);
        Assert.Contains("\"count\":0", result.Replace(" ", ""));
        Assert.Contains("No bookmarks found", result);
    }

    [Fact]
    public void DeleteBookmark_ShouldDeleteBookmark()
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
    public void DeleteBookmark_WithKeepTextFalse_ShouldRemoveBookmarkAndText()
    {
        var docPath = CreateWordDocument("test_delete_no_keep.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("DeleteWithText");
        builder.Write("TextToRemove");
        builder.EndBookmark("DeleteWithText");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_no_keep_output.docx");
        var result = _tool.Execute("delete", docPath, outputPath: outputPath, name: "DeleteWithText", keepText: false);
        Assert.Contains("deleted successfully", result);
    }

    [Fact]
    public void EditBookmark_ShouldEditBookmarkText()
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
    public void EditBookmark_WithNewName_ShouldRenameBookmark()
    {
        var docPath = CreateWordDocument("test_edit_name.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("OldName");
        builder.Write("Content");
        builder.EndBookmark("OldName");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_name_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, name: "OldName", newName: "NewName");
        var resultDoc = new Document(outputPath);
        Assert.Null(resultDoc.Range.Bookmarks["OldName"]);
        Assert.NotNull(resultDoc.Range.Bookmarks["NewName"]);
    }

    [Fact]
    public void EditBookmark_WithNewText_ShouldUpdateContent()
    {
        var docPath = CreateWordDocument("test_edit_text.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("EditText");
        builder.Write("Old content");
        builder.EndBookmark("EditText");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_text_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, name: "EditText", newText: "New content");
        var resultDoc = new Document(outputPath);
        Assert.Contains("New content", resultDoc.Range.Bookmarks["EditText"].Text);
    }

    [Fact]
    public void EditBookmark_WithBothNewNameAndNewText_ShouldUpdateBoth()
    {
        var docPath = CreateWordDocument("test_edit_both.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("OldBoth");
        builder.Write("Old");
        builder.EndBookmark("OldBoth");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_both_output.docx");
        var result = _tool.Execute("edit", docPath, outputPath: outputPath, name: "OldBoth", newName: "NewBoth",
            newText: "New");
        Assert.Contains("edited successfully", result);
    }

    [Fact]
    public void GotoBookmark_ShouldNavigateToBookmark()
    {
        var docPath = CreateWordDocumentWithContent("test_goto_bookmark.docx", "Content before bookmark");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Content after");
        builder.StartBookmark("TargetBookmark");
        builder.Write("Target content");
        builder.EndBookmark("TargetBookmark");
        doc.Save(docPath);

        var result = _tool.Execute("goto", docPath, name: "TargetBookmark");
        Assert.Contains("TargetBookmark", result);
        Assert.Contains("Target content", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_add_{operation}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, name: $"BM_{operation}", text: "Text");
        Assert.StartsWith("Bookmark added successfully", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_get_{operation}.docx", "Test");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var docPath = CreateWordDocument($"test_case_del_{operation}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark($"BM_{operation}");
        builder.EndBookmark($"BM_{operation}");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, name: $"BM_{operation}");
        Assert.Contains("deleted successfully", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var docPath = CreateWordDocument($"test_case_edit_{operation}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark($"BM_{operation}");
        builder.Write("Old");
        builder.EndBookmark($"BM_{operation}");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, name: $"BM_{operation}", newText: "New");
        Assert.Contains("edited successfully", result);
    }

    [Theory]
    [InlineData("GOTO")]
    [InlineData("Goto")]
    [InlineData("goto")]
    public void Operation_ShouldBeCaseInsensitive_Goto(string operation)
    {
        var docPath = CreateWordDocument($"test_case_goto_{operation}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark($"BM_{operation}");
        builder.Write("Content");
        builder.EndBookmark($"BM_{operation}");
        doc.Save(docPath);

        var result = _tool.Execute(operation, docPath, name: $"BM_{operation}");
        Assert.StartsWith("Bookmark location information", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void AddBookmark_WithEmptyOrNullName_ShouldThrowArgumentException(string? name)
    {
        var docPath = CreateWordDocumentWithContent("test_add_empty_name.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, name: name, text: "Some text"));
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
            _tool.Execute("add", docPath, outputPath: outputPath, name: "NewBookmark", text: "Text",
                paragraphIndex: 999));
        Assert.Contains("out of range", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void EditBookmark_WithEmptyOrNullName_ShouldThrowArgumentException(string? name)
    {
        var docPath = CreateWordDocumentWithContent("test_edit_empty_name.docx", "Test content");
        var outputPath = CreateTestFilePath("test_edit_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, name: name, text: "New text"));
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
    }

    [Fact]
    public void EditBookmark_WithNoChanges_ShouldThrowArgumentException()
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

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void DeleteBookmark_WithEmptyOrNullName_ShouldThrowArgumentException(string? name)
    {
        var docPath = CreateWordDocumentWithContent("test_delete_empty_name.docx", "Test content");
        var outputPath = CreateTestFilePath("test_delete_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, name: name));
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
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void GotoBookmark_WithEmptyOrNullName_ShouldThrowArgumentException(string? name)
    {
        var docPath = CreateWordDocumentWithContent("test_goto_empty_name.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("goto", docPath, name: name));
        Assert.Contains("Bookmark name is required", ex.Message);
    }

    [Fact]
    public void GotoBookmark_WithNonExistentBookmark_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_goto_nonexistent.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("goto", docPath, name: "NonExistentBookmark"));
        Assert.Contains("not found", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void AddBookmark_WithSessionId_ShouldAddBookmarkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add.docx", "Test content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, name: "SessionBookmark", text: "Session text");
        Assert.Contains("SessionBookmark", result);

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
        var result = _tool.Execute("get", docPath1, sessionId);
        Assert.Contains("SessionBookmark", result);
        Assert.DoesNotContain("PathBookmark", result);
    }

    #endregion
}