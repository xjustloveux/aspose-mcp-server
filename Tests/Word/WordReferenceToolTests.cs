using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordReferenceToolTests : WordTestBase
{
    private readonly WordReferenceTool _tool = new();

    [Fact]
    public async Task AddTableOfContents_ShouldAddTOC()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_toc.docx", "Heading 1\nContent\nHeading 2\nMore content");
        var doc = new Document(docPath);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        if (paragraphs.Count > 0) paragraphs[0].ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        if (paragraphs.Count > 2) paragraphs[2].ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_toc_output.docx");
        var arguments = CreateArguments("add_table_of_contents", docPath, outputPath);
        arguments["title"] = "Table of Contents";
        arguments["maxLevel"] = 3;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.NotNull(resultDoc);
    }

    [Fact]
    public async Task UpdateTableOfContents_ShouldUpdateTOC()
    {
        // Arrange
        var docPath = CreateWordDocument("test_update_toc.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("New Heading");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_toc_output.docx");
        var arguments = CreateArguments("update_table_of_contents", docPath, outputPath);

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public async Task AddCrossReference_WithReferenceType_ShouldUseReferenceType()
    {
        // Arrange
        var docPath = CreateWordDocument("test_cross_reference.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Chapter 1");
        builder.StartBookmark("Chapter1");
        builder.Writeln("Content of Chapter 1");
        builder.EndBookmark("Chapter1");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_cross_reference_output.docx");
        var arguments = CreateArguments("add_cross_reference", docPath, outputPath);
        arguments["referenceType"] = "Bookmark";
        arguments["targetName"] = "Chapter1";
        arguments["referenceText"] = "See ";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Bookmark", result, StringComparison.OrdinalIgnoreCase);
    }
}