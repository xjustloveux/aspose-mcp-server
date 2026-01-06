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

    #region General

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
    public void GetLayouts_ShouldReturnLayoutsWithType()
    {
        var pptPath = CreateTestPresentation("test_get_layouts.pptx");
        var result = _tool.Execute("get_layouts", pptPath);
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
    public void Set_ShouldSetSlideLayout()
    {
        var pptPath = CreateTestPresentation("test_set_layout.pptx");
        var outputPath = CreateTestFilePath("test_set_layout_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, layout: "Blank", outputPath: outputPath);
        Assert.StartsWith("Layout 'Blank' set for slide 0", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[0].LayoutSlide.LayoutType);
    }

    [Fact]
    public void ApplyMaster_ShouldApplyToAllSlides()
    {
        var pptPath = CreateTestPresentation("test_apply_master.pptx", 3);
        var outputPath = CreateTestFilePath("test_apply_master_output.pptx");
        var result = _tool.Execute("apply_master", pptPath, masterIndex: 0, layoutIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Master", result);
        Assert.Contains("3 slides", result);
    }

    [Fact]
    public void ApplyMaster_WithSlideIndices_ShouldApplyToSpecificSlides()
    {
        var pptPath = CreateTestPresentation("test_apply_master_specific.pptx", 5);
        var outputPath = CreateTestFilePath("test_apply_master_specific_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 2, 4 });
        var result = _tool.Execute("apply_master", pptPath, masterIndex: 0, layoutIndex: 0,
            slideIndices: slideIndicesJson, outputPath: outputPath);
        Assert.StartsWith("Master", result);
        Assert.Contains("3 slides", result);
    }

    [Fact]
    public void ApplyLayoutRange_ShouldApplyToMultipleSlides()
    {
        var pptPath = CreateTestPresentation("test_apply_range.pptx", 5);
        var outputPath = CreateTestFilePath("test_apply_range_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 1, 2 });
        var result = _tool.Execute("apply_layout_range", pptPath, layout: "Blank", slideIndices: slideIndicesJson,
            outputPath: outputPath);
        Assert.StartsWith("Layout 'Blank' applied to", result);
        Assert.Contains("3 slide(s)", result);
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
        Assert.StartsWith("Theme applied", result);
        Assert.Contains("master(s) copied", result);
        using var resultPres = new Presentation(outputPath);
        Assert.True(resultPres.Masters.Count > originalMasterCount);
    }

    [Theory]
    [InlineData("GET_MASTERS")]
    [InlineData("Get_Masters")]
    [InlineData("get_masters")]
    public void Operation_ShouldBeCaseInsensitive_GetMasters(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_masters_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath);
        Assert.Contains("\"count\"", result);
    }

    [Theory]
    [InlineData("GET_LAYOUTS")]
    [InlineData("Get_Layouts")]
    [InlineData("get_layouts")]
    public void Operation_ShouldBeCaseInsensitive_GetLayouts(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_layouts_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath);
        Assert.Contains("mastersCount", result);
    }

    [Theory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive_Set(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_set_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, layout: "Blank", outputPath: outputPath);
        Assert.StartsWith("Layout 'Blank' set for slide 0", result);
    }

    [Theory]
    [InlineData("APPLY_MASTER")]
    [InlineData("Apply_Master")]
    [InlineData("apply_master")]
    public void Operation_ShouldBeCaseInsensitive_ApplyMaster(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_apply_master_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_apply_master_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, masterIndex: 0, layoutIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Master", result);
    }

    [Theory]
    [InlineData("APPLY_LAYOUT_RANGE")]
    [InlineData("Apply_Layout_Range")]
    [InlineData("apply_layout_range")]
    public void Operation_ShouldBeCaseInsensitive_ApplyLayoutRange(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_range_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_range_{operation.Replace("_", "")}_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute(operation, pptPath, layout: "Blank", slideIndices: slideIndicesJson,
            outputPath: outputPath);
        Assert.StartsWith("Layout 'Blank' applied to", result);
    }

    [Theory]
    [InlineData("APPLY_THEME")]
    [InlineData("Apply_Theme")]
    [InlineData("apply_theme")]
    public void Operation_ShouldBeCaseInsensitive_ApplyTheme(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_theme_{operation.Replace("_", "")}.pptx");
        var themePath = CreateThemePresentation($"theme_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_theme_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, themePath: themePath, outputPath: outputPath);
        Assert.StartsWith("Theme applied", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Set_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, layout: "Blank"));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Set_WithoutLayout_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_no_layout.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, slideIndex: 0));
        Assert.Contains("layout is required", ex.Message);
    }

    [Fact]
    public void Set_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, slideIndex: 99, layout: "Title"));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Set_WithUnsupportedLayoutType_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_unsupported.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 0, layout: "InvalidLayoutType"));
        Assert.Contains("Unknown layout type", ex.Message);
        Assert.Contains("Supported types", ex.Message);
    }

    [Fact]
    public void GetLayouts_WithInvalidMasterIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_get_layouts_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_layouts", pptPath, masterIndex: 99));
        Assert.Contains("master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ApplyMaster_WithoutMasterIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_master_no_master.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("apply_master", pptPath, layoutIndex: 0));
        Assert.Contains("masterIndex is required", ex.Message);
    }

    [Fact]
    public void ApplyMaster_WithoutLayoutIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_master_no_layout.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("apply_master", pptPath, masterIndex: 0));
        Assert.Contains("layoutIndex is required", ex.Message);
    }

    [Fact]
    public void ApplyMaster_WithInvalidMasterIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_master_invalid_master.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_master", pptPath, masterIndex: 99, layoutIndex: 0));
        Assert.Contains("master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ApplyMaster_WithInvalidLayoutIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_master_invalid_layout.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_master", pptPath, masterIndex: 0, layoutIndex: 99));
        Assert.Contains("layout", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ApplyMaster_WithInvalidSlideIndices_ShouldThrowArgumentException()
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
    public void ApplyLayoutRange_WithoutLayout_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_range_no_layout.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_layout_range", pptPath, slideIndices: slideIndicesJson));
        Assert.Contains("layout is required", ex.Message);
    }

    [Fact]
    public void ApplyLayoutRange_WithoutSlideIndices_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_range_no_indices.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("apply_layout_range", pptPath, layout: "Blank"));
        Assert.Contains("slideIndices is required", ex.Message);
    }

    [Fact]
    public void ApplyLayoutRange_WithInvalidSlideIndices_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_range_invalid.pptx", 3);
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 99 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply_layout_range", pptPath, layout: "Blank", slideIndices: slideIndicesJson));
        Assert.Contains("Invalid slide indices", ex.Message);
    }

    [Fact]
    public void ApplyLayoutRange_WithUnsupportedLayout_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_range_unsupported.pptx", 3);
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("apply_layout_range", pptPath,
            layout: "NonExistentLayout", slideIndices: slideIndicesJson));
        Assert.Contains("Unknown layout type", ex.Message);
    }

    [Fact]
    public void ApplyTheme_WithoutThemePath_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_apply_theme_no_path.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("apply_theme", pptPath));
        Assert.Contains("themePath is required", ex.Message);
    }

    [Fact]
    public void ApplyTheme_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var pptPath = CreateTestPresentation("test_apply_theme_notfound.pptx");
        var ex = Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("apply_theme", pptPath, themePath: "nonexistent_theme.potx"));
        Assert.Contains("Theme file not found", ex.Message);
    }

    #endregion

    #region Session

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
    public void GetLayouts_WithSessionId_ShouldReturnLayouts()
    {
        var pptPath = CreateTestPresentation("test_session_get_layouts.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_layouts", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("mastersCount", out _));
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public void Set_WithSessionId_ShouldSetLayoutInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set_layout.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, layout: "Blank");
        Assert.StartsWith("Layout 'Blank' set for slide 0", result);
        Assert.Contains("session", result);
        Assert.Equal(SlideLayoutType.Blank, ppt.Slides[0].LayoutSlide.LayoutType);
    }

    [Fact]
    public void ApplyMaster_WithSessionId_ShouldApplyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_apply_master.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("apply_master", sessionId: sessionId, masterIndex: 0, layoutIndex: 0);
        Assert.StartsWith("Master", result);
        Assert.Contains("3 slides", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void ApplyLayoutRange_WithSessionId_ShouldApplyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_apply_range.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 1 });
        var result = _tool.Execute("apply_layout_range", sessionId: sessionId, layout: "Blank",
            slideIndices: slideIndicesJson);
        Assert.StartsWith("Layout 'Blank' applied to", result);
        Assert.Contains("2 slide(s)", result);
        Assert.Contains("session", result);
        Assert.Equal(SlideLayoutType.Blank, ppt.Slides[0].LayoutSlide.LayoutType);
        Assert.Equal(SlideLayoutType.Blank, ppt.Slides[1].LayoutSlide.LayoutType);
    }

    [Fact]
    public void ApplyTheme_WithSessionId_ShouldApplyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_apply_theme.pptx");
        var themePath = CreateThemePresentation("session_theme.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var originalMasterCount = ppt.Masters.Count;
        var result = _tool.Execute("apply_theme", sessionId: sessionId, themePath: themePath);
        Assert.StartsWith("Theme applied", result);
        Assert.Contains("session", result);
        Assert.True(ppt.Masters.Count > originalMasterCount);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_masters", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestPresentation("test_path_layout.pptx");
        var pptPath2 = CreateTestPresentation("test_session_layout.pptx", 5);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("apply_master", pptPath1, sessionId, masterIndex: 0, layoutIndex: 0);
        Assert.StartsWith("Master", result);
        Assert.Contains("5 slides", result);
        Assert.Contains("session", result);
    }

    #endregion
}