using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptLayoutToolTests : TestBase
{
    private readonly PptLayoutTool _tool;

    public PptLayoutToolTests()
    {
        _tool = new PptLayoutTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreateThemePresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void GetMasters_ShouldReturnMasterSlides()
    {
        var pptPath = CreateTestPresentation("test_get_masters.pptx");
        var result = _tool.Execute("get_masters", pptPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public void Set_ShouldSetSlideLayout()
    {
        var pptPath = CreateTestPresentation("test_set_layout.pptx");
        var outputPath = CreateTestFilePath("test_set_layout_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, layout: "Blank", outputPath: outputPath);
        Assert.Contains("Layout 'Blank' set for slide 0", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[0].LayoutSlide.LayoutType);
    }

    [Fact]
    public void GetLayouts_ShouldReturnLayoutsWithType()
    {
        var pptPath = CreateTestPresentation("test_get_layouts.pptx");
        var result = _tool.Execute("get_layouts", pptPath);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("mastersCount", out _));
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public void GetLayouts_WithMasterIndex_ShouldReturnSpecificMasterLayouts()
    {
        var pptPath = CreateTestPresentation("test_get_layouts_master.pptx");
        var result = _tool.Execute("get_layouts", pptPath, masterIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("masterIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void ApplyMaster_ShouldApplyToAllSlides()
    {
        var pptPath = CreateTestPresentation("test_apply_master.pptx", 3);
        var outputPath = CreateTestFilePath("test_apply_master_output.pptx");
        var result = _tool.Execute("apply_master", pptPath, masterIndex: 0, layoutIndex: 0, outputPath: outputPath);
        Assert.Contains("applied to 3 slides", result);
    }

    [Fact]
    public void ApplyMaster_WithSlideIndices_ShouldApplyToSpecificSlides()
    {
        var pptPath = CreateTestPresentation("test_apply_master_specific.pptx", 5);
        var outputPath = CreateTestFilePath("test_apply_master_specific_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 2, 4 });
        var result = _tool.Execute("apply_master", pptPath, masterIndex: 0, layoutIndex: 0,
            slideIndices: slideIndicesJson, outputPath: outputPath);
        Assert.Contains("applied to 3 slides", result);
    }

    [Fact]
    public void ApplyLayoutRange_ShouldApplyToMultipleSlides()
    {
        var pptPath = CreateTestPresentation("test_apply_range.pptx", 5);
        var outputPath = CreateTestFilePath("test_apply_range_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 1, 2 });
        var result = _tool.Execute("apply_layout_range", pptPath, layout: "Blank", slideIndices: slideIndicesJson,
            outputPath: outputPath);
        Assert.Contains("applied to 3 slide(s)", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[0].LayoutSlide.LayoutType);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[1].LayoutSlide.LayoutType);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[2].LayoutSlide.LayoutType);
    }

    [Fact]
    public void ApplyTheme_ShouldCopyMasterSlides()
    {
        var pptPath = CreateTestPresentation("test_apply_theme.pptx", 3);
        var themePath = CreateThemePresentation("theme.pptx");
        var outputPath = CreateTestFilePath("test_apply_theme_output.pptx");

        using var origPres = new Presentation(pptPath);
        var originalMasterCount = origPres.Masters.Count;
        var result = _tool.Execute("apply_theme", pptPath, themePath: themePath, outputPath: outputPath);
        Assert.Contains("master(s) copied", result);
        Assert.Contains("layout applied to all slides", result);

        using var resultPres = new Presentation(outputPath);
        Assert.True(resultPres.Masters.Count > originalMasterCount);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Set_WithInvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_set_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, slideIndex: 99, layout: "Title"));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    [Fact]
    public void Set_WithUnsupportedLayoutType_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_set_unsupported.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 0, layout: "InvalidLayoutType"));
        Assert.Contains("Unknown layout type", ex.Message);
        Assert.Contains("Supported types", ex.Message);
    }

    [Fact]
    public void GetLayouts_WithInvalidMasterIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_get_layouts_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_layouts", pptPath, masterIndex: 99));
        Assert.Contains("master", ex.Message.ToLower());
    }

    [Fact]
    public void ApplyMaster_WithInvalidSlideIndices_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_apply_master_invalid.pptx", 3);
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 10, 20 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_master", pptPath, masterIndex: 0, layoutIndex: 0, slideIndices: slideIndicesJson));
        Assert.Contains("Invalid slide indices", ex.Message);
        Assert.Contains("10", ex.Message);
        Assert.Contains("20", ex.Message);
    }

    [Fact]
    public void ApplyMaster_WithInvalidMasterIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_apply_master_invalid_master.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_master", pptPath, masterIndex: 99, layoutIndex: 0));
        Assert.Contains("master", ex.Message.ToLower());
    }

    [Fact]
    public void ApplyLayoutRange_WithInvalidSlideIndices_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_apply_range_invalid.pptx", 3);
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 99 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_layout_range", pptPath, layout: "Blank", slideIndices: slideIndicesJson));
        Assert.Contains("Invalid slide indices", ex.Message);
    }

    [Fact]
    public void ApplyLayoutRange_WithUnsupportedLayout_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_apply_range_unsupported.pptx", 3);
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("apply_layout_range", pptPath,
            layout: "NonExistentLayout", slideIndices: slideIndicesJson));
        Assert.Contains("Unknown layout type", ex.Message);
    }

    [Fact]
    public void ApplyTheme_WithNonExistentFile_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_apply_theme_notfound.pptx");
        var ex = Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("apply_theme", pptPath, themePath: "nonexistent_theme.potx"));
        Assert.Contains("Theme file not found", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetMasters_WithSessionId_ShouldReturnMasterSlides()
    {
        var pptPath = CreateTestPresentation("test_session_get_masters.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_masters", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public void Set_WithSessionId_ShouldSetLayoutInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set_layout.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, layout: "Blank");
        Assert.Contains("Layout 'Blank' set for slide 0", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        Assert.Equal(SlideLayoutType.Blank, ppt.Slides[0].LayoutSlide.LayoutType);
    }

    [Fact]
    public void GetLayouts_WithSessionId_ShouldReturnLayouts()
    {
        var pptPath = CreateTestPresentation("test_session_get_layouts.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_layouts", sessionId: sessionId);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("mastersCount", out _));
        Assert.Contains("layoutType", result);
    }

    #endregion
}