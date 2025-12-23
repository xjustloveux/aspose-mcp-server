using System.Text.Json.Nodes;
using Aspose.Words;
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
}