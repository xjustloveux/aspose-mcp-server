using System.Drawing;
using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptBackgroundToolTests : TestBase
{
    private readonly PptBackgroundTool _tool;

    public PptBackgroundToolTests()
    {
        _tool = new PptBackgroundTool(SessionManager);
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

    #region General Tests

    #region Set Background Tests

    [Fact]
    public void SetBackgroundColor_ShouldSetBackgroundColor()
    {
        var pptPath = CreateTestPresentation("test_set_background_color.pptx");
        var outputPath = CreateTestFilePath("test_set_background_color_output.pptx");

        _tool.Execute("set", pptPath, slideIndex: 0, color: "#FF0000", outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Solid, presentation.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void SetBackground_ApplyToAll_ShouldApplyToAllSlides()
    {
        var pptPath = CreateTestPresentation("test_set_background_all.pptx", 3);
        var outputPath = CreateTestFilePath("test_set_background_all_output.pptx");

        var result = _tool.Execute("set", pptPath, color: "#00FF00", applyToAll: true, outputPath: outputPath);

        Assert.Contains("all", result);
        using var presentation = new Presentation(outputPath);
        foreach (var slide in presentation.Slides)
            Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
    }

    [Fact]
    public void SetBackgroundImage_ShouldSetPictureFill()
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
        File.WriteAllBytes(imagePath, pngBytes);

        _tool.Execute("set", pptPath, slideIndex: 0, imagePath: imagePath, outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Picture, presentation.Slides[0].Background.FillFormat.FillType);
    }

    #endregion

    #region Get Background Tests

    [Fact]
    public void GetBackground_ShouldReturnBackgroundInfo()
    {
        var pptPath = CreateTestPresentation("test_get_background.pptx");

        var result = _tool.Execute("get", pptPath, slideIndex: 0);

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("hasBackground", out _));
        Assert.True(json.RootElement.TryGetProperty("fillType", out _));
    }

    [Fact]
    public void GetBackground_WithSolidColor_ShouldReturnColorAndOpacity()
    {
        var pptPath = CreateTestPresentation("test_get_background_solid.pptx");

        using (var ppt = new Presentation(pptPath))
        {
            ppt.Slides[0].Background.Type = BackgroundType.OwnBackground;
            ppt.Slides[0].Background.FillFormat.FillType = FillType.Solid;
            ppt.Slides[0].Background.FillFormat.SolidFillColor.Color = Color.Red;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath, slideIndex: 0);

        var json = JsonDocument.Parse(result);
        Assert.Equal("Solid", json.RootElement.GetProperty("fillType").GetString());
        Assert.True(json.RootElement.TryGetProperty("color", out _));
        Assert.True(json.RootElement.TryGetProperty("opacity", out _));
    }

    #endregion

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_UnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_operation.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
    }

    [Fact]
    public void SetBackground_NoColorOrImage_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_background_no_params.pptx");
        var outputPath = CreateTestFilePath("test_set_background_no_params_output.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, slideIndex: 0, outputPath: outputPath));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetBackground_WithSessionId_ShouldReturnBackgroundInfo()
    {
        var pptPath = CreateTestPresentation("test_session_get_background.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("fillType", out _));
    }

    [Fact]
    public void SetBackground_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set_background.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, color: "#FF0000");
        Assert.Contains("Background", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(FillType.Solid, ppt.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void SetBackground_WithSessionId_ApplyToAll_ShouldApplyToAllSlides()
    {
        var pptPath = CreateTestPresentation("test_session_set_background_all.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, color: "#00FF00", applyToAll: true);
        Assert.Contains("all", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        foreach (var slide in ppt.Slides)
            Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
    }

    #endregion
}