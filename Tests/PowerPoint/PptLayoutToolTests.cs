using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptLayoutToolTests : TestBase
{
    private readonly PptLayoutTool _tool = new();

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

    #region GetMasters Operation Tests

    [Fact]
    public async Task GetMasters_ShouldReturnMasterSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_masters.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_masters",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
        Assert.Contains("layoutType", result);
    }

    #endregion

    #region Unknown Operation Tests

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Set Operation Tests

    [Fact]
    public async Task Set_ShouldSetSlideLayout()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_layout.pptx");
        var outputPath = CreateTestFilePath("test_set_layout_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["layout"] = "Blank"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Layout 'Blank' set for slide 0", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[0].LayoutSlide.LayoutType);
    }

    [Fact]
    public async Task Set_WithInvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_invalid_slide.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["slideIndex"] = 99,
            ["layout"] = "Title"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    [Fact]
    public async Task Set_WithUnsupportedLayoutType_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_unsupported.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["layout"] = "InvalidLayoutType"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown layout type", ex.Message);
        Assert.Contains("Supported types", ex.Message);
    }

    #endregion

    #region GetLayouts Operation Tests

    [Fact]
    public async Task GetLayouts_ShouldReturnLayoutsWithType()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_layouts.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_layouts",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("mastersCount", out _));
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public async Task GetLayouts_WithMasterIndex_ShouldReturnSpecificMasterLayouts()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_layouts_master.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_layouts",
            ["path"] = pptPath,
            ["masterIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("masterIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public async Task GetLayouts_WithInvalidMasterIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_layouts_invalid.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_layouts",
            ["path"] = pptPath,
            ["masterIndex"] = 99
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("master", ex.Message.ToLower());
    }

    #endregion

    #region ApplyMaster Operation Tests

    [Fact]
    public async Task ApplyMaster_ShouldApplyToAllSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_master.pptx", 3);
        var outputPath = CreateTestFilePath("test_apply_master_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply_master",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["masterIndex"] = 0,
            ["layoutIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("applied to 3 slides", result);
    }

    [Fact]
    public async Task ApplyMaster_WithSlideIndices_ShouldApplyToSpecificSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_master_specific.pptx", 5);
        var outputPath = CreateTestFilePath("test_apply_master_specific_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply_master",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["masterIndex"] = 0,
            ["layoutIndex"] = 0,
            ["slideIndices"] = new JsonArray { 0, 2, 4 }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("applied to 3 slides", result);
    }

    [Fact]
    public async Task ApplyMaster_WithInvalidSlideIndices_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_master_invalid.pptx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "apply_master",
            ["path"] = pptPath,
            ["masterIndex"] = 0,
            ["layoutIndex"] = 0,
            ["slideIndices"] = new JsonArray { 0, 10, 20 }
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid slide indices", ex.Message);
        Assert.Contains("10", ex.Message);
        Assert.Contains("20", ex.Message);
    }

    [Fact]
    public async Task ApplyMaster_WithInvalidMasterIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_master_invalid_master.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply_master",
            ["path"] = pptPath,
            ["masterIndex"] = 99,
            ["layoutIndex"] = 0
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("master", ex.Message.ToLower());
    }

    #endregion

    #region ApplyLayoutRange Operation Tests

    [Fact]
    public async Task ApplyLayoutRange_ShouldApplyToMultipleSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_range.pptx", 5);
        var outputPath = CreateTestFilePath("test_apply_range_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply_layout_range",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["layout"] = "Blank",
            ["slideIndices"] = new JsonArray { 0, 1, 2 }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("applied to 3 slide(s)", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[0].LayoutSlide.LayoutType);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[1].LayoutSlide.LayoutType);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[2].LayoutSlide.LayoutType);
    }

    [Fact]
    public async Task ApplyLayoutRange_WithInvalidSlideIndices_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_range_invalid.pptx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "apply_layout_range",
            ["path"] = pptPath,
            ["layout"] = "Blank",
            ["slideIndices"] = new JsonArray { 0, 99 }
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid slide indices", ex.Message);
    }

    [Fact]
    public async Task ApplyLayoutRange_WithUnsupportedLayout_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_range_unsupported.pptx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "apply_layout_range",
            ["path"] = pptPath,
            ["layout"] = "NonExistentLayout",
            ["slideIndices"] = new JsonArray { 0 }
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown layout type", ex.Message);
    }

    #endregion

    #region ApplyTheme Operation Tests

    [Fact]
    public async Task ApplyTheme_ShouldCopyMasterSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_theme.pptx", 3);
        var themePath = CreateThemePresentation("theme.pptx");
        var outputPath = CreateTestFilePath("test_apply_theme_output.pptx");

        using var origPres = new Presentation(pptPath);
        var originalMasterCount = origPres.Masters.Count;

        var arguments = new JsonObject
        {
            ["operation"] = "apply_theme",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["themePath"] = themePath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("master(s) copied", result);
        Assert.Contains("layout applied to all slides", result);

        using var resultPres = new Presentation(outputPath);
        Assert.True(resultPres.Masters.Count > originalMasterCount);
    }

    [Fact]
    public async Task ApplyTheme_WithNonExistentFile_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_apply_theme_notfound.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply_theme",
            ["path"] = pptPath,
            ["themePath"] = "nonexistent_theme.potx"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Theme file not found", ex.Message);
    }

    #endregion
}