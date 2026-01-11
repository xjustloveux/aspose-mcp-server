using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordParagraphTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordParagraphToolTests : WordTestBase
{
    private readonly WordParagraphTool _tool;

    public WordParagraphToolTests()
    {
        _tool = new WordParagraphTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Insert_ShouldInsertParagraphAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_insert.docx", "First", "Third");
        var outputPath = CreateTestFilePath("test_insert_output.docx");
        _tool.Execute("insert", docPath, outputPath: outputPath, paragraphIndex: 0, text: "Second");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.Contains(paragraphs, p => p.GetText().Contains("Second"));
    }

    [Fact]
    public void Get_ShouldReturnParagraphsFromFile()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_get.docx", "First", "Second", "Third");
        var result = _tool.Execute("get", docPath);
        Assert.Contains("First", result);
        Assert.Contains("Second", result);
        Assert.Contains("Third", result);
    }

    [Fact]
    public void Delete_ShouldDeleteParagraphAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_delete.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_delete_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, paragraphIndex: 1);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc, false);
        Assert.DoesNotContain(paragraphs, p => p.GetText().Trim().Contains("Second"));
    }

    [Fact]
    public void Edit_ShouldEditParagraphAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_edit.docx", "Test content");
        var outputPath = CreateTestFilePath("test_edit_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, paragraphIndex: 0, bold: true);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GetFormat_ShouldReturnFormatInfoFromFile()
    {
        var docPath = CreateWordDocumentWithContent("test_get_format.docx", "Formatted text");
        var result = _tool.Execute("get_format", docPath, paragraphIndex: 0);
        Assert.Contains("alignment", result);
    }

    [Fact]
    public void Merge_ShouldMergeParagraphsAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_merge.docx", "First", "Second", "Third");
        var outputPath = CreateTestFilePath("test_merge_output.docx");
        var result = _tool.Execute("merge", docPath, outputPath: outputPath, startParagraphIndex: 0,
            endParagraphIndex: 2);
        Assert.StartsWith("Paragraphs merged", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("INSERT")]
    [InlineData("Insert")]
    [InlineData("insert")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        _tool.Execute(operation, docPath, outputPath: outputPath, text: "Test paragraph");
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Content");
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
    public void Insert_WithSessionId_ShouldInsertInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_insert.docx", "Existing content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("insert", sessionId: sessionId, text: "New paragraph", paragraphIndex: 0);
        Assert.StartsWith("Paragraph inserted successfully", result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = GetParagraphs(doc);
        Assert.Contains(paragraphs, p => p.GetText().Contains("New paragraph"));
    }

    [Fact]
    public void Get_WithSessionId_ShouldReturnParagraphInfo()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_get.docx", "First", "Second", "Third");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("First", result);
        Assert.Contains("Second", result);
        Assert.Contains("Third", result);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_delete.docx", "First", "Second", "Third");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("delete", sessionId: sessionId, paragraphIndex: 1);
        Assert.StartsWith("Paragraph #1 deleted", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_edit.docx", "Test content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("edit", sessionId: sessionId, paragraphIndex: 0, bold: true, fontSize: 16);
        Assert.Contains("format edited successfully", result);
    }

    [Fact]
    public void GetFormat_WithSessionId_ShouldReturnFormat()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_format.docx", "Formatted text");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_format", sessionId: sessionId, paragraphIndex: 0);
        Assert.Contains("alignment", result);
    }

    [Fact]
    public void CopyFormat_WithSessionId_ShouldCopyInMemory()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_copy_format.docx", "Source", "Target");
        var sessionId = OpenSession(docPath);
        var doc = SessionManager.GetDocument<Document>(sessionId);

        var paragraphs = GetParagraphs(doc);
        paragraphs[0].ParagraphFormat.LeftIndent = 72;
        _tool.Execute("copy_format", sessionId: sessionId, sourceParagraphIndex: 0, targetParagraphIndex: 1);
        var resultParagraphs = GetParagraphs(doc);
        Assert.Equal(72, resultParagraphs[1].ParagraphFormat.LeftIndent);
    }

    [Fact]
    public void Merge_WithSessionId_ShouldMergeInMemory()
    {
        var docPath = CreateWordDocumentWithParagraphs("test_session_merge.docx", "First", "Second", "Third");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("merge", sessionId: sessionId, startParagraphIndex: 0, endParagraphIndex: 2);
        Assert.StartsWith("Paragraphs merged", result);
        Assert.Contains("session", result);
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
        var docPath1 = CreateWordDocumentWithContent("test_path_para.docx", "Path content");
        var docPath2 = CreateWordDocumentWithContent("test_session_para.docx", "Session content");
        var sessionId = OpenSession(docPath2);

        var result = _tool.Execute("get", docPath1, sessionId);

        Assert.Contains("Session content", result);
    }

    #endregion
}
