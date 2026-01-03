using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptSectionToolTests : TestBase
{
    private readonly PptSectionTool _tool;

    public PptSectionToolTests()
    {
        _tool = new PptSectionTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddSection_ShouldAddSection()
    {
        var pptPath = CreateTestPresentation("test_add_section.pptx");
        var outputPath = CreateTestFilePath("test_add_section_output.pptx");
        _tool.Execute("add", pptPath, name: "Section 1", slideIndex: 0, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Sections.Count > 0, "Presentation should contain at least one section");
    }

    [Fact]
    public void GetSections_ShouldReturnAllSectionsWithSlideInfo()
    {
        var pptPath = CreateTestPresentation("test_get_sections.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            presentation.Sections.AddSection("Test Section", presentation.Slides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Test Section", result);
        Assert.Contains("startSlideIndex", result);
        Assert.Contains("slideCount", result);
    }

    [Fact]
    public void GetSections_WhenNoSections_ShouldReturnEmptyResult()
    {
        var pptPath = CreateTestPresentation("test_get_sections_empty.pptx");
        var result = _tool.Execute("get", pptPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No sections found", result);
    }

    [Fact]
    public void RenameSection_ShouldRenameSection()
    {
        var pptPath = CreateTestPresentation("test_rename_section.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Old Name", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_rename_section_output.pptx");
        _tool.Execute("rename", pptPath, sectionIndex: 0, newName: "New Name", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Equal("New Name", presentation.Sections[0].Name);
    }

    [Fact]
    public void DeleteSection_ShouldDeleteSectionKeepSlides()
    {
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
        _tool.Execute("delete", pptPath, sectionIndex: 0, keepSlides: true, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Sections.Count < sectionsBefore, "Section should be deleted");
        Assert.Equal(slidesBefore, presentation.Slides.Count);
    }

    [Fact]
    public void DeleteSection_WithKeepSlidesFalse_ShouldDeleteSectionAndSlides()
    {
        var pptPath = CreateTestPresentation("test_delete_section_with_slides.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Section to Delete", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_section_with_slides_output.pptx");
        _tool.Execute("delete", pptPath, sectionIndex: 0, keepSlides: false, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Sections);
    }

    [Fact]
    public void RenameSection_WithInvalidIndex_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_rename_invalid.pptx");
        using (var ppt = new Presentation(pptPath))
        {
            ppt.Sections.AddSection("Test", ppt.Slides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        Assert.Throws<ArgumentException>(() => _tool.Execute("rename", pptPath, sectionIndex: 99, newName: "New Name"));
    }

    [Fact]
    public void DeleteSection_WithInvalidIndex_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_delete_invalid.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("delete", pptPath, sectionIndex: 99));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_WithUnknownOperation_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddSection_MissingName_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_add_missing_name.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0));
        Assert.Contains("name", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetSections_WithSessionId_ShouldReturnSectionsFromMemory()
    {
        var pptPath = CreateTestPresentation("test_session_get.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            presentation.Sections.AddSection("Session Section", presentation.Slides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("Session Section", result);
    }

    [Fact]
    public void AddSection_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Sections.Count;
        var result = _tool.Execute("add", sessionId: sessionId, name: "New Session Section", slideIndex: 0);
        Assert.Contains("Section", result);
        Assert.True(ppt.Sections.Count > initialCount);
    }

    [Fact]
    public void RenameSection_WithSessionId_ShouldRenameInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_rename.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            presentation.Sections.AddSection("Old Section Name", presentation.Slides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("rename", sessionId: sessionId, sectionIndex: 0, newName: "Renamed Section");
        Assert.Contains("renamed", result, StringComparison.OrdinalIgnoreCase);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal("Renamed Section", ppt.Sections[0].Name);
    }

    [Fact]
    public void DeleteSection_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_delete.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            presentation.Sections.AddSection("Section To Delete", presentation.Slides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Sections.Count;
        var result = _tool.Execute("delete", sessionId: sessionId, sectionIndex: 0, keepSlides: true);
        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(ppt.Sections.Count < initialCount);
    }

    #endregion
}