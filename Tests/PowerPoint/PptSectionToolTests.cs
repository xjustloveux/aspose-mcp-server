using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptSectionToolTests : TestBase
{
    private readonly PptSectionTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task AddSection_ShouldAddSection()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_section.pptx");
        var outputPath = CreateTestFilePath("test_add_section_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["name"] = "Section 1",
            ["slideIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Sections.Count > 0, "Presentation should contain at least one section");
    }

    [Fact]
    public async Task GetSections_ShouldReturnAllSections()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_sections.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            presentation.Sections.AddSection("Test Section", presentation.Slides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Section", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RenameSection_ShouldRenameSection()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_rename_section.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Old Name", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_rename_section_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "rename",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["sectionIndex"] = 0,
            ["newName"] = "New Name"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.Equal("New Name", presentation.Sections[0].Name);
    }

    [Fact]
    public async Task DeleteSection_ShouldDeleteSection()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_section.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Section to Delete", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var sectionsBefore = new Presentation(pptPath).Sections.Count;
        Assert.True(sectionsBefore > 0, "Section should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_section_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["sectionIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var sectionsAfter = presentation.Sections.Count;
        Assert.True(sectionsAfter < sectionsBefore,
            $"Section should be deleted. Before: {sectionsBefore}, After: {sectionsAfter}");
    }
}