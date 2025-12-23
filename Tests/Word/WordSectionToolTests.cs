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
        Assert.Contains("Section", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteSection_ShouldDeleteSection()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        // Insert a section break to create a new section
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        doc.Save(docPath);

        // Reload to ensure sections are properly created
        doc = new Document(docPath);
        var sectionsBefore = doc.Sections.Count;
        Assert.True(sectionsBefore > 1, "Document should have multiple sections before deletion");

        var outputPath = CreateTestFilePath("test_delete_section_output.docx");
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["sectionIndex"] = 1;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var sectionsAfter = resultDoc.Sections.Count;
        Assert.True(sectionsAfter < sectionsBefore,
            $"Section should be deleted. Before: {sectionsBefore}, After: {sectionsAfter}");
    }
}