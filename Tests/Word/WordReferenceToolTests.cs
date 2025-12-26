using System.Text.Json.Nodes;
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

    [Fact]
    public async Task AddTableOfContents_AtEnd_ShouldInsertAtEnd()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_toc_end.docx", "Content before TOC");
        var outputPath = CreateTestFilePath("test_toc_end_output.docx");
        var arguments = CreateArguments("add_table_of_contents", docPath, outputPath);
        arguments["position"] = "end";
        arguments["title"] = "Contents";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Table of contents added", result);
    }

    [Fact]
    public async Task UpdateTableOfContents_WhenNoTOC_ShouldReturnMessage()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_no_toc.docx", "Document without TOC");
        var arguments = CreateArguments("update_table_of_contents", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("No table of contents fields found", result);
        Assert.Contains("add_table_of_contents", result);
    }

    [Fact]
    public async Task UpdateTableOfContents_WithInvalidTocIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_toc_invalid_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\"");
        doc.Save(docPath);

        var arguments = CreateArguments("update_table_of_contents", docPath);
        arguments["tocIndex"] = 5;

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("tocIndex must be between", exception.Message);
    }

    [Fact]
    public async Task AddIndex_ShouldAddIndexEntries()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_index.docx", "Document content");
        var outputPath = CreateTestFilePath("test_add_index_output.docx");
        var arguments = CreateArguments("add_index", docPath, outputPath);
        arguments["indexEntries"] = new JsonArray
        {
            new JsonObject { ["text"] = "Term1" },
            new JsonObject { ["text"] = "Term2", ["subEntry"] = "SubTerm" }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Index entries added", result);
        Assert.Contains("Total entries: 2", result);
    }

    [Fact]
    public async Task AddIndex_WithInvalidHeadingStyle_ShouldFallbackToHeading1()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_index_invalid_style.docx", "Content");
        var outputPath = CreateTestFilePath("test_index_invalid_style_output.docx");
        var arguments = CreateArguments("add_index", docPath, outputPath);
        arguments["indexEntries"] = new JsonArray
        {
            new JsonObject { ["text"] = "TestTerm" }
        };
        arguments["headingStyle"] = "NonExistentStyle";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Index entries added", result);
    }

    [Fact]
    public async Task AddCrossReference_WithInvalidReferenceType_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_invalid_ref_type.docx");
        var arguments = CreateArguments("add_cross_reference", docPath);
        arguments["referenceType"] = "InvalidType";
        arguments["targetName"] = "SomeTarget";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid referenceType", exception.Message);
        Assert.Contains("Heading, Bookmark, Figure, Table, Equation", exception.Message);
    }

    [Fact]
    public async Task AddCrossReference_WithIncludeAboveBelow_ShouldAddText()
    {
        // Arrange
        var docPath = CreateWordDocument("test_cross_ref_above.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("Target");
        builder.Writeln("Target content");
        builder.EndBookmark("Target");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_cross_ref_above_output.docx");
        var arguments = CreateArguments("add_cross_reference", docPath, outputPath);
        arguments["referenceType"] = "Bookmark";
        arguments["targetName"] = "Target";
        arguments["includeAboveBelow"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Cross-reference added", result);
    }

    [Fact]
    public async Task AddIndex_WithoutInsertingIndexField_ShouldOnlyAddEntries()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_index_no_field.docx", "Content");
        var outputPath = CreateTestFilePath("test_index_no_field_output.docx");
        var arguments = CreateArguments("add_index", docPath, outputPath);
        arguments["indexEntries"] = new JsonArray
        {
            new JsonObject { ["text"] = "Entry1" }
        };
        arguments["insertIndexAtEnd"] = false;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Index entries added", result);
    }
}