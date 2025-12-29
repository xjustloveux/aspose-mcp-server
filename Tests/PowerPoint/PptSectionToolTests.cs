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
    public async Task GetSections_ShouldReturnAllSectionsWithSlideInfo()
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
        Assert.Contains("Test Section", result);
        Assert.Contains("startSlideIndex", result);
        Assert.Contains("slideCount", result);
    }

    [Fact]
    public async Task GetSections_WhenNoSections_ShouldReturnEmptyResult()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_sections_empty.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No sections found", result);
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
    public async Task DeleteSection_ShouldDeleteSectionKeepSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_section.pptx");
        int slidesBefore;
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Section to Delete", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
            slidesBefore = ppt.Slides.Count;
        }

        int sectionsBefore;
        using (var pptCheck = new Presentation(pptPath))
        {
            sectionsBefore = pptCheck.Sections.Count;
        }

        Assert.True(sectionsBefore > 0, "Section should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_section_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["sectionIndex"] = 0,
            ["keepSlides"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Sections.Count < sectionsBefore, "Section should be deleted");
        Assert.Equal(slidesBefore, presentation.Slides.Count);
    }

    [Fact]
    public async Task DeleteSection_WithKeepSlidesFalse_ShouldDeleteSectionAndSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_section_with_slides.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Section to Delete", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_section_with_slides_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["sectionIndex"] = 0,
            ["keepSlides"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Sections);
    }

    [Fact]
    public async Task RenameSection_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_rename_invalid.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Test", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "rename",
            ["path"] = pptPath,
            ["sectionIndex"] = 99,
            ["newName"] = "New Name"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task DeleteSection_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_invalid.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["sectionIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ExecuteAsync_WithUnknownOperation_ShouldThrowException()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}