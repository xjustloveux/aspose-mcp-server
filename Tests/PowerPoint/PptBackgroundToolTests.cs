using System.Drawing;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptBackgroundToolTests : TestBase
{
    private readonly PptBackgroundTool _tool = new();

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region Error Handling Tests

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_operation.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region Set Background Tests

    [Fact]
    public async Task SetBackgroundColor_ShouldSetBackgroundColor()
    {
        var pptPath = CreateTestPresentation("test_set_background_color.pptx");
        var outputPath = CreateTestFilePath("test_set_background_color_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["color"] = "#FF0000"
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Solid, presentation.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public async Task SetBackground_ApplyToAll_ShouldApplyToAllSlides()
    {
        var pptPath = CreateTestPresentation("test_set_background_all.pptx", 3);
        var outputPath = CreateTestFilePath("test_set_background_all_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["color"] = "#00FF00",
            ["applyToAll"] = true
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("all", result);
        using var presentation = new Presentation(outputPath);
        foreach (var slide in presentation.Slides)
            Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
    }

    [Fact]
    public async Task SetBackground_NoColorOrImage_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_set_background_no_params.pptx");
        var outputPath = CreateTestFilePath("test_set_background_no_params_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SetBackgroundImage_ShouldSetPictureFill()
    {
        var pptPath = CreateTestPresentation("test_set_background_image.pptx");
        var outputPath = CreateTestFilePath("test_set_background_image_output.pptx");
        var imagePath = CreateTestFilePath("test_bg_image.png");

        // Create a minimal valid PNG file (1x1 pixel blue)
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 pixel
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0x0F, 0x00, 0x00,
            0x01, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB4,
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
            0xAE, 0x42, 0x60, 0x82 // IEND chunk
        };
        await File.WriteAllBytesAsync(imagePath, pngBytes);

        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imagePath"] = imagePath
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Picture, presentation.Slides[0].Background.FillFormat.FillType);
    }

    #endregion

    #region Get Background Tests

    [Fact]
    public async Task GetBackground_ShouldReturnBackgroundInfo()
    {
        var pptPath = CreateTestPresentation("test_get_background.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("hasBackground", out _));
        Assert.True(json.RootElement.TryGetProperty("fillType", out _));
    }

    [Fact]
    public async Task GetBackground_WithSolidColor_ShouldReturnColorAndOpacity()
    {
        var pptPath = CreateTestPresentation("test_get_background_solid.pptx");

        using (var ppt = new Presentation(pptPath))
        {
            ppt.Slides[0].Background.Type = BackgroundType.OwnBackground;
            ppt.Slides[0].Background.FillFormat.FillType = FillType.Solid;
            ppt.Slides[0].Background.FillFormat.SolidFillColor.Color = Color.Red;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.Equal("Solid", json.RootElement.GetProperty("fillType").GetString());
        Assert.True(json.RootElement.TryGetProperty("color", out _));
        Assert.True(json.RootElement.TryGetProperty("opacity", out _));
    }

    #endregion
}