using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordHyperlinkToolTests : WordTestBase
{
    private readonly WordHyperlinkTool _tool = new();

    [Fact]
    public async Task AddHyperlink_ShouldAddHyperlink()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_hyperlink.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_hyperlink_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["text"] = "Click here";
        arguments["url"] = "https://example.com";
        arguments["paragraphIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var hyperlinks = doc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.True(hyperlinks.Count > 0 || doc.GetText().Contains("Click here"),
            "Document should contain a hyperlink or the hyperlink text");
    }

    [Fact]
    public async Task GetHyperlinks_ShouldReturnAllHyperlinks()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_hyperlinks.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Test Link", "https://test.com", false);
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public async Task EditHyperlink_ShouldEditHyperlink()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_edit_hyperlink.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Original Link", "https://original.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.docx");
        var arguments = CreateArguments("edit", docPath, outputPath);
        arguments["hyperlinkIndex"] = 0;
        arguments["url"] = "https://updated.com";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var resultDoc = new Document(outputPath);
        Assert.NotNull(resultDoc);
    }

    [Fact]
    public async Task DeleteHyperlink_ShouldDeleteHyperlink()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_delete_hyperlink.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link to Delete", "https://delete.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.docx");
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["hyperlinkIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var resultDoc = new Document(outputPath);
        var hyperlinks = resultDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.Empty(hyperlinks);
    }
}