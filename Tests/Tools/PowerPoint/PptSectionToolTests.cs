using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptSectionTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptSectionToolTests : PptTestBase
{
    private readonly PptSectionTool _tool;

    public PptSectionToolTests()
    {
        _tool = new PptSectionTool(SessionManager);
    }

    private string CreatePresentationWithSection(string fileName, string sectionName = "Test Section")
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Sections.AddSection(sectionName, presentation.Slides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddSection()
    {
        var pptPath = CreatePresentation("test_add.pptx");
        var outputPath = CreateTestFilePath("test_add_output.pptx");
        var result = _tool.Execute("add", pptPath, name: "Section 1", slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Section 'Section 1' added", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Sections.Count > 0);
        Assert.Equal("Section 1", presentation.Sections[0].Name);
    }

    [Fact]
    public void Get_ShouldReturnAllSections()
    {
        var pptPath = CreatePresentationWithSection("test_get.pptx");
        var result = _tool.Execute("get", pptPath);
        Assert.Contains("Test Section", result);
        Assert.Contains("\"count\": 1", result);
    }

    [Fact]
    public void Rename_ShouldRenameSection()
    {
        var pptPath = CreatePresentationWithSection("test_rename.pptx", "Old Name");
        var outputPath = CreateTestFilePath("test_rename_output.pptx");
        var result = _tool.Execute("rename", pptPath, sectionIndex: 0, newName: "New Name", outputPath: outputPath);
        Assert.StartsWith("Section 0 renamed to", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal("New Name", presentation.Sections[0].Name);
    }

    [Fact]
    public void Delete_WithKeepSlidesTrue_ShouldDeleteSectionKeepSlides()
    {
        var pptPath = CreatePresentationWithSection("test_delete_keep.pptx");
        var outputPath = CreateTestFilePath("test_delete_keep_output.pptx");
        var result = _tool.Execute("delete", pptPath, sectionIndex: 0, keepSlides: true, outputPath: outputPath);
        Assert.StartsWith("Section 0 removed", result);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Sections);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, name: "Section", slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Section 'Section' added", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldReturnSectionsFromMemory()
    {
        var pptPath = CreatePresentationWithSection("test_session_get.pptx", "Session Section");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("Session Section", result);
        Assert.Contains("\"count\": 1", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Sections.Count;
        var result = _tool.Execute("add", sessionId: sessionId, name: "New Section", slideIndex: 0);
        Assert.StartsWith("Section 'New Section' added", result);
        Assert.True(ppt.Sections.Count > initialCount);
    }

    [Fact]
    public void Rename_WithSessionId_ShouldRenameInMemory()
    {
        var pptPath = CreatePresentationWithSection("test_session_rename.pptx", "Old Name");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("rename", sessionId: sessionId, sectionIndex: 0, newName: "Renamed Section");
        Assert.StartsWith("Section 0 renamed to", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal("Renamed Section", ppt.Sections[0].Name);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithSection("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Sections.Count;
        var result = _tool.Execute("delete", sessionId: sessionId, sectionIndex: 0, keepSlides: true);
        Assert.StartsWith("Section 0 removed", result);
        Assert.True(ppt.Sections.Count < initialCount);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithSection("test_path_section.pptx", "PathSection");
        var pptPath2 = CreatePresentationWithSection("test_session_section.pptx", "SessionSection");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId);
        Assert.Contains("SessionSection", result);
        Assert.DoesNotContain("PathSection", result);
    }

    #endregion
}
