using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordSectionToolTests : WordTestBase
{
    private readonly WordSectionTool _tool = new();

    [Fact]
    public async Task InsertSection_ShouldInsertSection()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_insert_section.docx", "Content before section");
        var outputPath = CreateTestFilePath("test_insert_section_output.docx");
        var arguments = CreateArguments("insert", docPath, outputPath);
        arguments["sectionBreakType"] = "NextPage";
        arguments["insertAtParagraphIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        Assert.True(doc.Sections.Count > 1, "Document should contain multiple sections");
    }

    [Fact]
    public async Task InsertSection_AtDocumentEnd_ShouldWork()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_insert_end.docx", "Content");
        var outputPath = CreateTestFilePath("test_insert_end_output.docx");
        var arguments = CreateArguments("insert", docPath, outputPath);
        arguments["sectionBreakType"] = "Continuous";
        arguments["insertAtParagraphIndex"] = -1; // Document end

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Section break inserted", result);
        Assert.Contains("Continuous", result);
    }

    [Fact]
    public async Task InsertSection_WithInvalidSectionIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_sec.docx", "Content");
        var arguments = CreateArguments("insert", docPath);
        arguments["sectionBreakType"] = "NextPage";
        arguments["sectionIndex"] = 99;
        arguments["insertAtParagraphIndex"] = 0;

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("sectionIndex must be between", exception.Message);
    }

    [Fact]
    public async Task InsertSection_WithInvalidParagraphIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_para.docx", "Content");
        var arguments = CreateArguments("insert", docPath);
        arguments["sectionBreakType"] = "NextPage";
        arguments["insertAtParagraphIndex"] = 999;

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("insertAtParagraphIndex must be between", exception.Message);
    }

    [Fact]
    public async Task GetSections_ShouldReturnSectionsInfo()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_sections.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.PageBreak);
        builder.CurrentSection.PageSetup.SectionStart = SectionStart.NewPage;
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("\"sections\"", result); // JSON format
        Assert.Contains("\"sectionBreak\"", result);
        Assert.Contains("\"type\"", result);
    }

    [Fact]
    public async Task GetSections_WithSpecificIndex_ShouldReturnSingleSection()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_single_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section content");
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);
        arguments["sectionIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"section\"", result); // Single section object
        Assert.Contains("\"index\": 0", result);
        Assert.DoesNotContain("\"index\": 1", result);
    }

    [Fact]
    public async Task GetSections_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_invalid_index.docx", "Content");
        var arguments = CreateArguments("get", docPath);
        arguments["sectionIndex"] = 99;

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("sectionIndex must be between", exception.Message);
    }

    [Fact]
    public async Task DeleteSection_ShouldDeleteSection()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        doc.Save(docPath);

        doc = new Document(docPath);
        var sectionsBefore = doc.Sections.Count;
        Assert.True(sectionsBefore > 1, "Document should have multiple sections before deletion");

        var outputPath = CreateTestFilePath("test_delete_section_output.docx");
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["sectionIndex"] = 1;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.Equal(sectionsBefore - 1, resultDoc.Sections.Count);
        Assert.Contains("Deleted", result);
        Assert.Contains("with their content", result);
    }

    [Fact]
    public async Task DeleteSection_LastSection_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_delete_last.docx", "Single section");
        var arguments = CreateArguments("delete", docPath);
        arguments["sectionIndex"] = 0;

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Cannot delete the last section", exception.Message);
    }

    [Fact]
    public async Task DeleteSection_MultipleSections_ShouldDeleteAll()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_multiple.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_multiple_output.docx");
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["sectionIndices"] = new JsonArray { 1, 2 };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        Assert.Equal(1, resultDoc.Sections.Count);
        Assert.Contains("Deleted 2 section(s)", result);
    }

    [Fact]
    public async Task DeleteSection_WithoutIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_no_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        doc.Save(docPath);

        var arguments = CreateArguments("delete", docPath);

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("sectionIndex or sectionIndices must be provided", exception.Message);
    }
}