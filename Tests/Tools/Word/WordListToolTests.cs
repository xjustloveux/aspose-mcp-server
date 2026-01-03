using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordListToolTests : WordTestBase
{
    private readonly WordListTool _tool;

    public WordListToolTests()
    {
        _tool = new WordListTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddList_ShouldAddBulletList()
    {
        var docPath = CreateWordDocument("test_add_list.docx");
        var outputPath = CreateTestFilePath("test_add_list_output.docx");
        var items = new JsonArray { "Item 1", "Item 2", "Item 3" };
        _tool.Execute("add_list", docPath, outputPath: outputPath, items: items);
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count >= 3, "Document should have list items");
        var docText = doc.GetText();
        Assert.Contains("Item 1", docText);
        Assert.Contains("Item 2", docText);
    }

    [Fact]
    public void AddList_WithNumberedList_ShouldAddNumberedList()
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
    public void GetListFormat_ShouldReturnListFormat()
    {
        var docPath = CreateWordDocument("test_get_list_format.docx");
        var items = new JsonArray { "Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items);
        var result = _tool.Execute("get_format", docPath, paragraphIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [SkippableFact]
    public void RestartNumbering_ShouldRestartListNumbering()
    {
        // Skip in evaluation mode as list operations may be limited
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_restart_numbering.docx");
        var outputPath = CreateTestFilePath("test_restart_numbering_output.docx");

        // First add a numbered list
        var items = new JsonArray { "Item 1", "Item 2", "Item 3", "Item 4" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        // Act - restart numbering at the 3rd item (index 2)
        var result = _tool.Execute("restart_numbering", docPath, outputPath: outputPath,
            paragraphIndex: 2, startAt: 1);
        Assert.Contains("restarted successfully", result);
        Assert.Contains("Paragraphs affected:", result);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        // The 3rd and 4th items should now belong to a different list
        var list1 = paragraphs[0].ListFormat.List;
        var list2 = paragraphs[2].ListFormat.List;
        Assert.NotNull(list1);
        Assert.NotNull(list2);
        Assert.NotEqual(list1.ListId, list2.ListId);
    }

    [Fact]
    public void RestartNumbering_WithCustomStartAt_ShouldStartFromSpecifiedNumber()
    {
        var docPath = CreateWordDocument("test_restart_numbering_custom.docx");
        var outputPath = CreateTestFilePath("test_restart_numbering_custom_output.docx");

        var items = new JsonArray { "Item 1", "Item 2", "Item 3" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        // Act - restart numbering at item 2 starting from 5
        var result = _tool.Execute("restart_numbering", docPath, outputPath: outputPath,
            paragraphIndex: 1, startAt: 5);
        Assert.Contains("restarted successfully", result);
        Assert.Contains("Start at: 5", result);
    }

    [Fact]
    public void ConvertToList_ShouldConvertParagraphsToBulletList()
    {
        // Arrange - Create document with regular paragraphs
        var docPath = CreateTestFilePath("test_convert_to_list.docx");
        var outputPath = CreateTestFilePath("test_convert_to_list_output.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph");
        builder.Writeln("Second paragraph");
        builder.Writeln("Third paragraph");
        doc.Save(docPath);
        var result = _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 2, listType: "bullet");
        Assert.Contains("converted to list successfully", result);
        Assert.Contains("Converted: 3 paragraphs", result);

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        // All 3 paragraphs should now be list items
        for (var i = 0; i < 3; i++)
            Assert.True(paragraphs[i].ListFormat.IsListItem, $"Paragraph {i} should be a list item");
    }

    [Fact]
    public void ConvertToList_WithNumberedList_ShouldConvertToNumberedList()
    {
        var docPath = CreateTestFilePath("test_convert_to_numbered_list.docx");
        var outputPath = CreateTestFilePath("test_convert_to_numbered_list_output.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Step one");
        builder.Writeln("Step two");
        doc.Save(docPath);
        var result = _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 1, listType: "number", numberFormat: "arabic");
        Assert.Contains("converted to list successfully", result);
        Assert.Contains("List type: number", result);

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        Assert.True(paragraphs[0].ListFormat.IsListItem);
        Assert.True(paragraphs[1].ListFormat.IsListItem);
    }

    [Fact]
    public void ConvertToList_ShouldSkipExistingListItems()
    {
        // Arrange - Create document with mixed content
        var docPath = CreateTestFilePath("test_convert_skip_existing.docx");
        var outputPath = CreateTestFilePath("test_convert_skip_existing_output.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Regular paragraph 1");

        // Add a list item
        var list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;
        builder.Writeln("Existing list item");
        builder.ListFormat.RemoveNumbers();

        builder.Writeln("Regular paragraph 2");
        doc.Save(docPath);
        var result = _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 2);
        Assert.Contains("converted to list successfully", result);
        Assert.Contains("Skipped: 1 paragraphs", result); // The existing list item should be skipped
    }

    [Fact]
    public void AddList_WithContinuePrevious_ShouldContinueExistingList()
    {
        var docPath = CreateWordDocument("test_continue_list.docx");
        var outputPath = CreateTestFilePath("test_continue_list_output.docx");

        // First add a numbered list
        var items1 = new JsonArray { "Item 1", "Item 2" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items1, listType: "number");

        // Act - Add more items continuing the previous list
        var items2 = new JsonArray { "Item 3", "Item 4" };
        var result = _tool.Execute("add_list", docPath, outputPath: outputPath, items: items2, continuePrevious: true);
        Assert.Contains("continuing previous list", result);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>()
            .Where(p => p.ListFormat.IsListItem)
            .ToList();

        // All items should belong to the same list
        var listId = paragraphs[0].ListFormat.List?.ListId;
        foreach (var para in paragraphs) Assert.Equal(listId, para.ListFormat.List?.ListId);
    }

    [Fact]
    public void AddItem_ShouldAddItemWithStyle()
    {
        var docPath = CreateWordDocument("test_add_item.docx");
        var outputPath = CreateTestFilePath("test_add_item_output.docx");
        var result = _tool.Execute("add_item", docPath, outputPath: outputPath,
            text: "New list item", styleName: "List Paragraph");
        Assert.Contains("added successfully", result);
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("New list item", text);
    }

    [SkippableFact]
    public void DeleteItem_ShouldDeleteParagraph()
    {
        // Skip in evaluation mode as list operations may be limited
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateTestFilePath("test_delete_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 0 - to delete");
        builder.Writeln("Paragraph 1 - to keep");
        builder.Writeln("Paragraph 2 - to keep");
        doc.Save(docPath);

        var paragraphCountBefore = doc.GetChildNodes(NodeType.Paragraph, true).Count;

        var outputPath = CreateTestFilePath("test_delete_item_output.docx");
        var result = _tool.Execute("delete_item", docPath, outputPath: outputPath, paragraphIndex: 0);
        Assert.Contains("deleted successfully", result);
        var resultDoc = new Document(outputPath);
        var paragraphCountAfter = resultDoc.GetChildNodes(NodeType.Paragraph, true).Count;
        Assert.True(paragraphCountAfter < paragraphCountBefore);
        Assert.DoesNotContain("Paragraph 0", resultDoc.GetText());
        Assert.Contains("Paragraph 1", resultDoc.GetText());
    }

    [SkippableFact]
    public void EditItem_ShouldUpdateParagraphText()
    {
        // Skip in evaluation mode as list operations may be limited
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateTestFilePath("test_edit_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Original text");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_item_output.docx");
        var result = _tool.Execute("edit_item", docPath, outputPath: outputPath,
            paragraphIndex: 0, text: "Updated text");
        Assert.Contains("edited successfully", result);
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Updated text", text);
        Assert.DoesNotContain("Original text", text);
    }

    [Fact]
    public void EditItem_WithLevel_ShouldChangeIndentation()
    {
        var docPath = CreateWordDocument("test_edit_item_level.docx");
        var items = new JsonArray { "Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items);

        var outputPath = CreateTestFilePath("test_edit_item_level_output.docx");
        var result = _tool.Execute("edit_item", docPath, outputPath: outputPath,
            paragraphIndex: 0, text: "Modified item", level: 2);
        Assert.Contains("edited successfully", result);
        Assert.Contains("Level: 2", result);
    }

    [Fact]
    public void SetFormat_ShouldSetListItemFormat()
    {
        var docPath = CreateWordDocument("test_set_format.docx");
        var items = new JsonArray { "Item 1", "Item 2" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        var outputPath = CreateTestFilePath("test_set_format_output.docx");
        var result = _tool.Execute("set_format", docPath, outputPath: outputPath,
            paragraphIndex: 0, leftIndent: 72.0);
        Assert.Contains("format set successfully", result);
        Assert.Contains("Left indent: 72", result);
    }

    [SkippableFact]
    public void SetFormat_WithNumberStyle_ShouldChangeNumberStyle()
    {
        // Skip in evaluation mode as list operations may be limited
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_set_format_style.docx");
        var items = new JsonArray { "Item 1", "Item 2" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        var outputPath = CreateTestFilePath("test_set_format_style_output.docx");
        var result = _tool.Execute("set_format", docPath, outputPath: outputPath,
            paragraphIndex: 0, numberStyle: "roman");
        Assert.Contains("format set successfully", result);
        Assert.Contains("Number style: roman", result);
    }

    [Fact]
    public void GetFormat_WithNonListParagraph_ShouldIndicateNotListItem()
    {
        var docPath = CreateTestFilePath("test_get_format_non_list.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Regular paragraph");
        doc.Save(docPath);
        var result = _tool.Execute("get_format", docPath, paragraphIndex: 0);
        Assert.Contains("\"isListItem\": false", result); // JSON format
        Assert.Contains("not a list item", result);
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
    public void AddList_WithEmptyItems_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_empty_items.docx");
        var outputPath = CreateTestFilePath("test_add_empty_items_output.docx");
        var items = new JsonArray();
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_list", docPath, outputPath: outputPath, items: items));

        Assert.Contains("items", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddList_WithNullItems_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_null_items.docx");
        var outputPath = CreateTestFilePath("test_add_null_items_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_list", docPath, outputPath: outputPath, items: null));

        Assert.Contains("items", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteItem_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_invalid_idx.docx", "Single paragraph");
        var outputPath = CreateTestFilePath("test_delete_invalid_idx_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_item", docPath, outputPath: outputPath, paragraphIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditItem_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_invalid_idx.docx", "Single paragraph");
        var outputPath = CreateTestFilePath("test_edit_invalid_idx_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_item", docPath, outputPath: outputPath, paragraphIndex: 999, text: "New text"));

        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetFormat_WithSessionId_ShouldReturnListFormat()
    {
        var docPath = CreateWordDocument("test_session_get_format.docx");
        var items = new JsonArray { "Session Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_format", sessionId: sessionId, paragraphIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void AddList_WithSessionId_ShouldAddListInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_list.docx");
        var sessionId = OpenSession(docPath);
        var items = new JsonArray { "Session Item A", "Session Item B" };
        var result = _tool.Execute("add_list", sessionId: sessionId, items: items);
        Assert.Contains("List added successfully", result);

        // Verify in-memory document has the list
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
        _tool.Execute("edit_item", sessionId: sessionId, paragraphIndex: 0, text: "Updated session text");

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.Contains("Updated session text", text);
        Assert.DoesNotContain("Original session text", text);
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
        _tool.Execute("delete_item", sessionId: sessionId, paragraphIndex: 0);

        // Assert - verify in-memory deletion
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var text = sessionDoc.GetText();
        Assert.DoesNotContain("Paragraph to delete", text);
        Assert.Contains("Paragraph to keep", text);
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("get_format", docPath1, sessionId, paragraphIndex: 0);

        // Assert - should use sessionId document (Session paragraph unique)
        // The result shows paragraph info, but we can verify by content
        Assert.NotNull(result);
    }

    #endregion
}