using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptDataOperationsToolTests : TestBase
{
    private readonly PptDataOperationsTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task GetStatistics_ShouldReturnStatistics()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_statistics.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_statistics",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("totalSlides", result);
    }

    [Fact]
    public async Task GetContent_ShouldReturnContent()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_content.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        // Content may be empty for blank slides, just verify result is returned
        Assert.True(result.Length > 0, "Result should not be empty");
    }

    [Fact]
    public async Task GetSlideDetails_ShouldReturnSlideDetails()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_slide_details.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_slide_details",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("slideIndex", result);
        Assert.Contains("slideSize", result);
    }

    [Fact]
    public async Task GetStatistics_ShouldIncludeHiddenSlidesCount()
    {
        var pptPath = CreateTestFilePath("test_hidden_slides.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Slides[0].Hidden = true;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_statistics",
            ["path"] = pptPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("totalHiddenSlides", result);
    }

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetSlideDetails_InvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_slide_details",
            ["path"] = pptPath,
            ["slideIndex"] = 99
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetContent_ShouldIncludeHiddenFlag()
    {
        var pptPath = CreateTestFilePath("test_content_hidden.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Slides[0].Hidden = true;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pptPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("hidden", result);
    }

    [Fact]
    public async Task GetSlideDetails_WithThumbnail_ShouldReturnBase64Image()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_thumbnail.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_slide_details",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["includeThumbnail"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("thumbnail", result);
        // Base64 PNG starts with iVBORw0KGgo (PNG signature)
        Assert.Contains("iVBORw0KGgo", result);
    }

    [Fact]
    public async Task GetSlideDetails_WithoutThumbnail_ShouldNotIncludeThumbnailData()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_no_thumbnail.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_slide_details",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["includeThumbnail"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("thumbnail", result);
        Assert.Contains("null", result); // thumbnail should be null
    }
}