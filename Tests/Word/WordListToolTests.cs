using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordListToolTests : WordTestBase
{
    private readonly WordListTool _tool = new();

    [Fact]
    public async Task AddList_ShouldAddBulletList()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_list.docx");
        var outputPath = CreateTestFilePath("test_add_list_output.docx");
        var items = new JsonArray { "Item 1", "Item 2", "Item 3" };
        var arguments = CreateArguments("add_list", docPath, outputPath);
        arguments["items"] = items;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count >= 3, "Document should have list items");
        var docText = doc.GetText();
        Assert.Contains("Item 1", docText);
        Assert.Contains("Item 2", docText);
    }

    [Fact]
    public async Task AddList_WithNumberedList_ShouldAddNumberedList()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_numbered_list.docx");
        var outputPath = CreateTestFilePath("test_add_numbered_list_output.docx");
        var items = new JsonArray { "First", "Second" };
        var arguments = CreateArguments("add_list", docPath, outputPath);
        arguments["items"] = items;
        arguments["listType"] = "number";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var docText = doc.GetText();
        Assert.Contains("First", docText);
        Assert.Contains("Second", docText);
    }

    [Fact]
    public async Task GetListFormat_ShouldReturnListFormat()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_list_format.docx");
        var items = new JsonArray { "Item 1" };
        var addArgs = CreateArguments("add_list", docPath, docPath);
        addArgs["items"] = items;
        await _tool.ExecuteAsync(addArgs);

        var arguments = CreateArguments("get_format", docPath);
        arguments["paragraphIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public async Task RestartNumbering_ShouldRestartListNumbering()
    {
        // Skip in evaluation mode as list operations may be limited
        if (IsEvaluationMode()) return;

        // Arrange
        var docPath = CreateWordDocument("test_restart_numbering.docx");
        var outputPath = CreateTestFilePath("test_restart_numbering_output.docx");

        // First add a numbered list
        var items = new JsonArray { "Item 1", "Item 2", "Item 3", "Item 4" };
        var addArgs = CreateArguments("add_list", docPath, docPath);
        addArgs["items"] = items;
        addArgs["listType"] = "number";
        await _tool.ExecuteAsync(addArgs);

        // Act - restart numbering at the 3rd item (index 2)
        var arguments = CreateArguments("restart_numbering", docPath, outputPath);
        arguments["paragraphIndex"] = 2;
        arguments["startAt"] = 1;

        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task RestartNumbering_WithCustomStartAt_ShouldStartFromSpecifiedNumber()
    {
        // Arrange
        var docPath = CreateWordDocument("test_restart_numbering_custom.docx");
        var outputPath = CreateTestFilePath("test_restart_numbering_custom_output.docx");

        var items = new JsonArray { "Item 1", "Item 2", "Item 3" };
        var addArgs = CreateArguments("add_list", docPath, docPath);
        addArgs["items"] = items;
        addArgs["listType"] = "number";
        await _tool.ExecuteAsync(addArgs);

        // Act - restart numbering at item 2 starting from 5
        var arguments = CreateArguments("restart_numbering", docPath, outputPath);
        arguments["paragraphIndex"] = 1;
        arguments["startAt"] = 5;

        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("restarted successfully", result);
        Assert.Contains("Start at: 5", result);
    }

    [Fact]
    public async Task ConvertToList_ShouldConvertParagraphsToBulletList()
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

        // Act
        var arguments = CreateArguments("convert_to_list", docPath, outputPath);
        arguments["startParagraphIndex"] = 0;
        arguments["endParagraphIndex"] = 2;
        arguments["listType"] = "bullet";

        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("converted to list successfully", result);
        Assert.Contains("Converted: 3 paragraphs", result);

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        // All 3 paragraphs should now be list items
        for (var i = 0; i < 3; i++)
            Assert.True(paragraphs[i].ListFormat.IsListItem, $"Paragraph {i} should be a list item");
    }

    [Fact]
    public async Task ConvertToList_WithNumberedList_ShouldConvertToNumberedList()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_convert_to_numbered_list.docx");
        var outputPath = CreateTestFilePath("test_convert_to_numbered_list_output.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Step one");
        builder.Writeln("Step two");
        doc.Save(docPath);

        // Act
        var arguments = CreateArguments("convert_to_list", docPath, outputPath);
        arguments["startParagraphIndex"] = 0;
        arguments["endParagraphIndex"] = 1;
        arguments["listType"] = "number";
        arguments["numberFormat"] = "arabic";

        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("converted to list successfully", result);
        Assert.Contains("List type: number", result);

        var resultDoc = new Document(outputPath);
        var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        Assert.True(paragraphs[0].ListFormat.IsListItem);
        Assert.True(paragraphs[1].ListFormat.IsListItem);
    }

    [Fact]
    public async Task ConvertToList_ShouldSkipExistingListItems()
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

        // Act
        var arguments = CreateArguments("convert_to_list", docPath, outputPath);
        arguments["startParagraphIndex"] = 0;
        arguments["endParagraphIndex"] = 2;

        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("converted to list successfully", result);
        Assert.Contains("Skipped: 1 paragraphs", result); // The existing list item should be skipped
    }

    [Fact]
    public async Task AddList_WithContinuePrevious_ShouldContinueExistingList()
    {
        // Arrange
        var docPath = CreateWordDocument("test_continue_list.docx");
        var outputPath = CreateTestFilePath("test_continue_list_output.docx");

        // First add a numbered list
        var items1 = new JsonArray { "Item 1", "Item 2" };
        var addArgs1 = CreateArguments("add_list", docPath, docPath);
        addArgs1["items"] = items1;
        addArgs1["listType"] = "number";
        await _tool.ExecuteAsync(addArgs1);

        // Act - Add more items continuing the previous list
        var items2 = new JsonArray { "Item 3", "Item 4" };
        var addArgs2 = CreateArguments("add_list", docPath, outputPath);
        addArgs2["items"] = items2;
        addArgs2["continuePrevious"] = true;

        var result = await _tool.ExecuteAsync(addArgs2);

        // Assert
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
    public async Task AddItem_ShouldAddItemWithStyle()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_item.docx");
        var outputPath = CreateTestFilePath("test_add_item_output.docx");
        var arguments = CreateArguments("add_item", docPath, outputPath);
        arguments["text"] = "New list item";
        arguments["styleName"] = "List Paragraph";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("added successfully", result);
        var doc = new Document(outputPath);
        var text = doc.GetText();
        Assert.Contains("New list item", text);
    }

    [Fact]
    public async Task DeleteItem_ShouldDeleteParagraph()
    {
        // Skip in evaluation mode as list operations may be limited
        if (IsEvaluationMode()) return;

        // Arrange
        var docPath = CreateTestFilePath("test_delete_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 0 - to delete");
        builder.Writeln("Paragraph 1 - to keep");
        builder.Writeln("Paragraph 2 - to keep");
        doc.Save(docPath);

        var paragraphCountBefore = doc.GetChildNodes(NodeType.Paragraph, true).Count;

        var outputPath = CreateTestFilePath("test_delete_item_output.docx");
        var arguments = CreateArguments("delete_item", docPath, outputPath);
        arguments["paragraphIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("deleted successfully", result);
        var resultDoc = new Document(outputPath);
        var paragraphCountAfter = resultDoc.GetChildNodes(NodeType.Paragraph, true).Count;
        Assert.True(paragraphCountAfter < paragraphCountBefore);
        Assert.DoesNotContain("Paragraph 0", resultDoc.GetText());
        Assert.Contains("Paragraph 1", resultDoc.GetText());
    }

    [Fact]
    public async Task EditItem_ShouldUpdateParagraphText()
    {
        // Skip in evaluation mode as list operations may be limited
        if (IsEvaluationMode()) return;

        // Arrange
        var docPath = CreateTestFilePath("test_edit_item.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Original text");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_item_output.docx");
        var arguments = CreateArguments("edit_item", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["text"] = "Updated text";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("edited successfully", result);
        var resultDoc = new Document(outputPath);
        var text = resultDoc.GetText();
        Assert.Contains("Updated text", text);
        Assert.DoesNotContain("Original text", text);
    }

    [Fact]
    public async Task EditItem_WithLevel_ShouldChangeIndentation()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_item_level.docx");
        var items = new JsonArray { "Item 1" };
        var addArgs = CreateArguments("add_list", docPath, docPath);
        addArgs["items"] = items;
        await _tool.ExecuteAsync(addArgs);

        var outputPath = CreateTestFilePath("test_edit_item_level_output.docx");
        var arguments = CreateArguments("edit_item", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["text"] = "Modified item";
        arguments["level"] = 2;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("edited successfully", result);
        Assert.Contains("Level: 2", result);
    }

    [Fact]
    public async Task SetFormat_ShouldSetListItemFormat()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_format.docx");
        var items = new JsonArray { "Item 1", "Item 2" };
        var addArgs = CreateArguments("add_list", docPath, docPath);
        addArgs["items"] = items;
        addArgs["listType"] = "number";
        await _tool.ExecuteAsync(addArgs);

        var outputPath = CreateTestFilePath("test_set_format_output.docx");
        var arguments = CreateArguments("set_format", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["leftIndent"] = 72.0; // 1 inch

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("format set successfully", result);
        Assert.Contains("Left indent: 72", result);
    }

    [Fact]
    public async Task SetFormat_WithNumberStyle_ShouldChangeNumberStyle()
    {
        // Skip in evaluation mode as list operations may be limited
        if (IsEvaluationMode()) return;

        // Arrange
        var docPath = CreateWordDocument("test_set_format_style.docx");
        var items = new JsonArray { "Item 1", "Item 2" };
        var addArgs = CreateArguments("add_list", docPath, docPath);
        addArgs["items"] = items;
        addArgs["listType"] = "number";
        await _tool.ExecuteAsync(addArgs);

        var outputPath = CreateTestFilePath("test_set_format_style_output.docx");
        var arguments = CreateArguments("set_format", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["numberStyle"] = "roman";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("format set successfully", result);
        Assert.Contains("Number style: roman", result);
    }

    [Fact]
    public async Task GetFormat_WithNonListParagraph_ShouldIndicateNotListItem()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_get_format_non_list.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Regular paragraph");
        doc.Save(docPath);

        var arguments = CreateArguments("get_format", docPath);
        arguments["paragraphIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"isListItem\": false", result); // JSON format
        Assert.Contains("not a list item", result);
    }
}