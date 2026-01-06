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

    #region General

    [Fact]
    public void AddList_ShouldAddBulletList()
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
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_restart_numbering.docx");
        var outputPath = CreateTestFilePath("test_restart_numbering_output.docx");

        var items = new JsonArray { "Item 1", "Item 2", "Item 3", "Item 4" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        var result = _tool.Execute("restart_numbering", docPath, outputPath: outputPath,
            paragraphIndex: 2, startAt: 1);
        Assert.StartsWith("List numbering restarted", result);
        Assert.Contains("Paragraphs affected:", result);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

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

        var result = _tool.Execute("restart_numbering", docPath, outputPath: outputPath,
            paragraphIndex: 1, startAt: 5);
        Assert.StartsWith("List numbering restarted", result);
        Assert.Contains("Start at: 5", result);
    }

    [Fact]
    public void ConvertToList_ShouldConvertParagraphsToBulletList()
    {
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
        Assert.StartsWith("Paragraphs converted to list", result);
        Assert.Contains("Converted: 3 paragraphs", result);

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

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
        Assert.StartsWith("Paragraphs converted to list", result);
        Assert.Contains("List type: number", result);

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        Assert.True(paragraphs[0].ListFormat.IsListItem);
        Assert.True(paragraphs[1].ListFormat.IsListItem);
    }

    [Fact]
    public void ConvertToList_ShouldSkipExistingListItems()
    {
        var docPath = CreateTestFilePath("test_convert_skip_existing.docx");
        var outputPath = CreateTestFilePath("test_convert_skip_existing_output.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Regular paragraph 1");

        var list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;
        builder.Writeln("Existing list item");
        builder.ListFormat.RemoveNumbers();

        builder.Writeln("Regular paragraph 2");
        doc.Save(docPath);
        var result = _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 2);
        Assert.StartsWith("Paragraphs converted to list", result);
        Assert.Contains("Skipped: 1 paragraphs", result);
    }

    [Fact]
    public void AddList_WithContinuePrevious_ShouldContinueExistingList()
    {
        var docPath = CreateWordDocument("test_continue_list.docx");
        var outputPath = CreateTestFilePath("test_continue_list_output.docx");

        var items1 = new JsonArray { "Item 1", "Item 2" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items1, listType: "number");

        var items2 = new JsonArray { "Item 3", "Item 4" };
        var result = _tool.Execute("add_list", docPath, outputPath: outputPath, items: items2, continuePrevious: true);
        Assert.Contains("continuing previous list", result);

        var doc = new Document(outputPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>()
            .Where(p => p.ListFormat.IsListItem)
            .ToList();

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
        Assert.StartsWith("List item added", result);
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("New list item", text);
    }

    [SkippableFact]
    public void DeleteItem_ShouldDeleteParagraph()
    {
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
        Assert.StartsWith("List item #0 deleted", result);
        var resultDoc = new Document(outputPath);
        var paragraphCountAfter = resultDoc.GetChildNodes(NodeType.Paragraph, true).Count;
        Assert.True(paragraphCountAfter < paragraphCountBefore);
        Assert.DoesNotContain("Paragraph 0", resultDoc.GetText());
        Assert.Contains("Paragraph 1", resultDoc.GetText());
    }

    [SkippableFact]
    public void EditItem_ShouldUpdateParagraphText()
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
        Assert.StartsWith("List item edited", result);
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
        Assert.StartsWith("List item edited", result);
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
        Assert.StartsWith("List format set", result);
        Assert.Contains("Left indent: 72", result);
    }

    [SkippableFact]
    public void SetFormat_WithNumberStyle_ShouldChangeNumberStyle()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "List operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_set_format_style.docx");
        var items = new JsonArray { "Item 1", "Item 2" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");

        var outputPath = CreateTestFilePath("test_set_format_style_output.docx");
        var result = _tool.Execute("set_format", docPath, outputPath: outputPath,
            paragraphIndex: 0, numberStyle: "roman");
        Assert.StartsWith("List format set", result);
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
        Assert.Contains("\"isListItem\": false", result);
        Assert.Contains("not a list item", result);
    }

    [Theory]
    [InlineData("ADD_LIST")]
    [InlineData("Add_List")]
    [InlineData("add_list")]
    public void Operation_ShouldBeCaseInsensitive_AddList(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var items = new JsonArray { "Item 1" };
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, items: items);
        Assert.StartsWith("List added", result);
    }

    [Theory]
    [InlineData("GET_FORMAT")]
    [InlineData("Get_Format")]
    [InlineData("get_format")]
    public void Operation_ShouldBeCaseInsensitive_GetFormat(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var items = new JsonArray { "Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items);
        var result = _tool.Execute(operation, docPath, paragraphIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Theory]
    [InlineData("ADD_ITEM")]
    [InlineData("Add_Item")]
    [InlineData("add_item")]
    public void Operation_ShouldBeCaseInsensitive_AddItem(string operation)
    {
        var docPath = CreateWordDocument($"test_case_item_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_item_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "New item", styleName: "List Paragraph");
        Assert.StartsWith("List item added", result);
    }

    [Theory]
    [InlineData("DELETE_ITEM")]
    [InlineData("Delete_Item")]
    [InlineData("delete_item")]
    public void Operation_ShouldBeCaseInsensitive_DeleteItem(string operation)
    {
        var docPath = CreateTestFilePath($"test_case_delete_{operation}.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph to delete");
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, paragraphIndex: 0);
        Assert.StartsWith("List item #0 deleted", result);
    }

    [Theory]
    [InlineData("EDIT_ITEM")]
    [InlineData("Edit_Item")]
    [InlineData("edit_item")]
    public void Operation_ShouldBeCaseInsensitive_EditItem(string operation)
    {
        var docPath = CreateTestFilePath($"test_case_edit_{operation}.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Original");
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, paragraphIndex: 0, text: "New");
        Assert.StartsWith("List item edited", result);
    }

    [Theory]
    [InlineData("SET_FORMAT")]
    [InlineData("Set_Format")]
    [InlineData("set_format")]
    public void Operation_ShouldBeCaseInsensitive_SetFormat(string operation)
    {
        var docPath = CreateWordDocument($"test_case_set_{operation}.docx");
        var items = new JsonArray { "Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items);
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, paragraphIndex: 0, leftIndent: 36.0);
        Assert.StartsWith("List format set", result);
    }

    [Theory]
    [InlineData("RESTART_NUMBERING")]
    [InlineData("Restart_Numbering")]
    [InlineData("restart_numbering")]
    public void Operation_ShouldBeCaseInsensitive_RestartNumbering(string operation)
    {
        var docPath = CreateWordDocument($"test_case_restart_{operation}.docx");
        var items = new JsonArray { "Item 1", "Item 2" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");
        var outputPath = CreateTestFilePath($"test_case_restart_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, paragraphIndex: 1, startAt: 1);
        Assert.StartsWith("List numbering restarted", result);
    }

    [Theory]
    [InlineData("CONVERT_TO_LIST")]
    [InlineData("Convert_To_List")]
    [InlineData("convert_to_list")]
    public void Operation_ShouldBeCaseInsensitive_ConvertToList(string operation)
    {
        var docPath = CreateTestFilePath($"test_case_convert_{operation}.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph");
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_convert_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            startParagraphIndex: 0, endParagraphIndex: 0);
        Assert.StartsWith("Paragraphs converted to list", result);
    }

    [Theory]
    [InlineData("BULLET")]
    [InlineData("Bullet")]
    [InlineData("bullet")]
    public void ListType_ShouldBeCaseInsensitive_Bullet(string listType)
    {
        var docPath = CreateWordDocument($"test_listtype_{listType}.docx");
        var outputPath = CreateTestFilePath($"test_listtype_{listType}_output.docx");
        var items = new JsonArray { "Item 1" };
        var result = _tool.Execute("add_list", docPath, outputPath: outputPath, items: items, listType: listType);
        Assert.StartsWith("List added", result);
    }

    [Theory]
    [InlineData("NUMBER")]
    [InlineData("Number")]
    [InlineData("number")]
    public void ListType_ShouldBeCaseInsensitive_Number(string listType)
    {
        var docPath = CreateWordDocument($"test_listtype_num_{listType}.docx");
        var outputPath = CreateTestFilePath($"test_listtype_num_{listType}_output.docx");
        var items = new JsonArray { "Item 1" };
        var result = _tool.Execute("add_list", docPath, outputPath: outputPath, items: items, listType: listType);
        Assert.StartsWith("List added", result);
        Assert.Contains("type: number", result.ToLower());
    }

    [Theory]
    [InlineData("ROMAN")]
    [InlineData("Roman")]
    [InlineData("roman")]
    public void NumberFormat_ShouldBeCaseInsensitive(string numberFormat)
    {
        var docPath = CreateWordDocument($"test_numformat_{numberFormat}.docx");
        var outputPath = CreateTestFilePath($"test_numformat_{numberFormat}_output.docx");
        var items = new JsonArray { "Item 1" };
        var result = _tool.Execute("add_list", docPath, outputPath: outputPath, items: items,
            listType: "number", numberFormat: numberFormat);
        Assert.StartsWith("List added", result);
        Assert.Contains("number format: roman", result.ToLower());
    }

    [Theory]
    [InlineData("ARABIC")]
    [InlineData("Arabic")]
    [InlineData("arabic")]
    public void NumberStyle_ShouldBeCaseInsensitive(string numberStyle)
    {
        var docPath = CreateWordDocument($"test_numstyle_{numberStyle}.docx");
        var items = new JsonArray { "Item 1" };
        _tool.Execute("add_list", docPath, outputPath: docPath, items: items, listType: "number");
        var outputPath = CreateTestFilePath($"test_numstyle_{numberStyle}_output.docx");
        var result = _tool.Execute("set_format", docPath, outputPath: outputPath,
            paragraphIndex: 0, numberStyle: numberStyle);
        Assert.StartsWith("List format set", result);
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

    [Fact]
    public void AddItem_WithMissingText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_item_no_text.docx");
        var outputPath = CreateTestFilePath("test_add_item_no_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_item", docPath, outputPath: outputPath, text: null, styleName: "List Paragraph"));

        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddItem_WithMissingStyleName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_item_no_style.docx");
        var outputPath = CreateTestFilePath("test_add_item_no_style_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_item", docPath, outputPath: outputPath, text: "Item", styleName: null));

        Assert.Contains("styleName", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetFormat_WithMissingParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_set_format_no_idx.docx");
        var outputPath = CreateTestFilePath("test_set_format_no_idx_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_format", docPath, outputPath: outputPath, paragraphIndex: null));

        Assert.Contains("paragraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConvertToList_WithMissingStartIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_convert_no_start.docx", "Paragraph");
        var outputPath = CreateTestFilePath("test_convert_no_start_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
                startParagraphIndex: null, endParagraphIndex: 0));

        Assert.Contains("startParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConvertToList_WithMissingEndIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_convert_no_end.docx", "Paragraph");
        var outputPath = CreateTestFilePath("test_convert_no_end_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
                startParagraphIndex: 0, endParagraphIndex: null));

        Assert.Contains("endParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConvertToList_WithStartGreaterThanEnd_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_convert_invalid_range.docx", "Paragraph");
        var outputPath = CreateTestFilePath("test_convert_invalid_range_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert_to_list", docPath, outputPath: outputPath,
                startParagraphIndex: 5, endParagraphIndex: 2));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void RestartNumbering_WithNonListParagraph_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_restart_nonlist.docx", "Regular paragraph");
        var outputPath = CreateTestFilePath("test_restart_nonlist_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("restart_numbering", docPath, outputPath: outputPath, paragraphIndex: 0));

        Assert.Contains("not a list item", ex.Message);
    }

    #endregion

    #region Session

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
        Assert.StartsWith("List added", result);

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

        var result = _tool.Execute("get_format", docPath1, sessionId, paragraphIndex: 0);

        Assert.NotNull(result);
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

        Assert.StartsWith("Paragraphs converted to list", result);

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

        Assert.StartsWith("List format set", result);
        Assert.Contains("Left indent: 50", result);
    }

    #endregion
}