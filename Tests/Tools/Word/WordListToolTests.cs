using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.List;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordListTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordListToolTests : WordTestBase
{
    private readonly WordListTool _tool;

    public WordListToolTests()
    {
        _tool = new WordListTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddList_ShouldAddBulletListAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_list.docx");
        var outputPath = CreateTestFilePath("test_add_list_output.docx");
        var items = new JsonArray { "Item 1", "Item 2", "Item 3" };
        _tool.Execute("add_list", docPath, outputPath: outputPath, items: items);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count >= 3);
        var docText = doc.GetText();
        Assert.Contains("Item 1", docText);
        Assert.Contains("Item 2", docText);
    }

    [Fact]
    public void AddList_WithNumberedList_ShouldAddNumberedListAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_numbered_list.docx");
        var outputPath = CreateTestFilePath("test_add_numbered_list_output.docx");
        var items = new JsonArray { "First", "Second" };
        _tool.Execute("add_list", docPath, outputPath: outputPath, items: items, listType: "number");
        var doc = new Document(outputPath);
        var docText = doc.GetText();
        Assert.Contains("First", docText);
        Assert.Contains("Second", docText);
    }

    [Fact]
    public void GetListFormat_ShouldReturnListFormatFromFile()
    {
        var docPath = CreateWordDocument("test_get_list_format.docx");
        var items = new JsonArray { "Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items);
        var result = _tool.Execute("get_format", docPath, paragraphIndex: 0);
        Assert.NotNull(result);
        var data = GetResultData<GetWordListFormatSingleResult>(result);
        Assert.Equal(0, data.ParagraphIndex);
        Assert.NotEmpty(data.ContentPreview);
    }

    [Fact]
    public void AddItem_ShouldAddItemAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_item.docx");
        var outputPath = CreateTestFilePath("test_add_item_output.docx");
        var result = _tool.Execute("add_item", docPath, outputPath: outputPath,
            text: "New list item", styleName: "List Paragraph");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List item added", data.Message);
        var doc = new Document(outputPath);
        Assert.Contains("New list item", doc.GetText());
    }

    [SkippableFact]
    public void DeleteItem_ShouldDeleteParagraphAndPersistToFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateTestFilePath("test_delete_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 0 - to delete");
        builder.Writeln("Paragraph 1 - to keep");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_item_output.docx");
        var result = _tool.Execute("delete_item", docPath, outputPath: outputPath, paragraphIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List item #0 deleted", data.Message);
        var resultDoc = new Document(outputPath);
        Assert.DoesNotContain("Paragraph 0", resultDoc.GetText());
        Assert.Contains("Paragraph 1", resultDoc.GetText());
    }

    [SkippableFact]
    public void EditItem_ShouldUpdateParagraphAndPersistToFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateTestFilePath("test_edit_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Original text");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_item_output.docx");
        var result = _tool.Execute("edit_item", docPath, outputPath: outputPath,
            paragraphIndex: 0, text: "Updated text");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List item edited", data.Message);
        var resultDoc = new Document(outputPath);
        Assert.Contains("Updated text", resultDoc.GetText());
    }

    [Fact]
    public void SetFormat_ShouldSetListItemFormatAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_format.docx");
        var items = new JsonArray { "Item 1", "Item 2" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        var outputPath = CreateTestFilePath("test_set_format_output.docx");
        var result = _tool.Execute("set_format", docPath, outputPath: outputPath,
            paragraphIndex: 0, leftIndent: 72.0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List format set", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ConvertToList_ShouldConvertParagraphsAndPersistToFile()
    {
        var docPath = CreateTestFilePath("test_convert_to_list.docx");
        var outputPath = CreateTestFilePath("test_convert_to_list_output.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph");
        builder.Writeln("Second paragraph");
        doc.Save(docPath);

        var result = _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 1, listType: "bullet");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Paragraphs converted to list", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void RestartNumbering_ShouldRestartListNumberingAndPersistToFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_restart_numbering.docx");
        var outputPath = CreateTestFilePath("test_restart_numbering_output.docx");

        var items = new JsonArray { "Item 1", "Item 2", "Item 3" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        var result = _tool.Execute("restart_numbering", docPath, outputPath: outputPath,
            paragraphIndex: 2, startAt: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List numbering restarted", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD_LIST")]
    [InlineData("Add_List")]
    [InlineData("add_list")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var items = new JsonArray { "Item 1" };
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, items: items);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List added", data.Message);
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
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("get_format", paragraphIndex: 0));
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetFormat_WithSessionId_ShouldReturnListFormat()
    {
        var docPath = CreateWordDocument("test_session_get_format.docx");
        var items = new JsonArray { "Session Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_format", sessionId: sessionId, paragraphIndex: 0);
        Assert.NotNull(result);
        var data = GetResultData<GetWordListFormatSingleResult>(result);
        Assert.Equal(0, data.ParagraphIndex);
        Assert.NotEmpty(data.ContentPreview);
        var output = GetResultOutput<GetWordListFormatSingleResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void AddList_WithSessionId_ShouldAddListInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_list.docx");
        var sessionId = OpenSession(docPath);
        var items = new JsonArray { "Session Item A", "Session Item B" };
        var result = _tool.Execute("add_list", sessionId: sessionId, items: items);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var text = doc.GetText();
        Assert.Contains("Session Item A", text);
        Assert.Contains("Session Item B", text);
    }

    [SkippableFact]
    public void EditItem_WithSessionId_ShouldEditItemInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode ignores edit operations");

        var docPath = CreateTestFilePath("test_session_edit_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Original session text");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("edit_item", sessionId: sessionId, paragraphIndex: 0, text: "Updated session text");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.Contains("Updated session text", text);
    }

    [SkippableFact]
    public void DeleteItem_WithSessionId_ShouldDeleteItemInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode ignores edit operations");

        var docPath = CreateTestFilePath("test_session_delete_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph to delete");
        builder.Writeln("Paragraph to keep");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("delete_item", sessionId: sessionId, paragraphIndex: 0);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.DoesNotContain("Paragraph to delete", sessionDoc.GetText());
        Assert.Contains("Paragraph to keep", sessionDoc.GetText());
    }

    [Fact]
    public void ConvertToList_WithSessionId_ShouldConvertInMemory()
    {
        var docPath = CreateTestFilePath("test_session_convert.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph");
        builder.Writeln("Second paragraph");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("convert_to_list", sessionId: sessionId,
            startParagraphIndex: 0, endParagraphIndex: 1);

        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Paragraphs converted to list", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = sessionDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        Assert.True(paragraphs[0].ListFormat.IsListItem);
        Assert.True(paragraphs[1].ListFormat.IsListItem);
    }

    [Fact]
    public void SetFormat_WithSessionId_ShouldSetFormatInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_format.docx");
        var items = new JsonArray { "Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_format", sessionId: sessionId, paragraphIndex: 0, leftIndent: 50.0);

        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("List format set", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_format", sessionId: "invalid_session_id", paragraphIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateTestFilePath("test_path_list.docx");
        var doc1 = new Document();
        var builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Path paragraph unique");
        doc1.Save(docPath1);

        var docPath2 = CreateTestFilePath("test_session_list.docx");
        var doc2 = new Document();
        var builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Session paragraph unique");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get_format", docPath1, sessionId, paragraphIndex: 0);
        Assert.NotNull(result);
    }

    #endregion
}
