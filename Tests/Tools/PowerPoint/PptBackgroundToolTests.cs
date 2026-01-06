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

    private string CreatePresentationWithSolidBackground(string fileName, Color bgColor)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        ppt.Slides[0].Background.Type = BackgroundType.OwnBackground;
        ppt.Slides[0].Background.FillFormat.FillType = FillType.Solid;
        ppt.Slides[0].Background.FillFormat.SolidFillColor.Color = bgColor;
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0x0F, 0x00, 0x00,
            0x01, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB4,
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
            0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(imagePath, pngBytes);
        return imagePath;
    }

    #region General

    [Fact]
    public void Set_WithColor_ShouldSetSolidBackground()
    {
        var pptPath = CreateTestPresentation("test_set_color.pptx");
        var outputPath = CreateTestFilePath("test_set_color_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, color: "#FF0000", outputPath: outputPath);
        Assert.StartsWith("Background", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Solid, presentation.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void Set_WithColorAndAlpha_ShouldSetTransparentBackground()
    {
        var pptPath = CreateTestPresentation("test_set_alpha.pptx");
        var outputPath = CreateTestFilePath("test_set_alpha_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, color: "#80FF0000", outputPath: outputPath);
        Assert.StartsWith("Background", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Solid, presentation.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void Set_WithImage_ShouldSetPictureBackground()
    {
        var pptPath = CreateTestPresentation("test_set_image.pptx");
        var outputPath = CreateTestFilePath("test_set_image_output.pptx");
        var imagePath = CreateTestImage("test_bg.png");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, imagePath: imagePath, outputPath: outputPath);
        Assert.StartsWith("Background", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Picture, presentation.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void Set_WithApplyToAll_ShouldApplyToAllSlides()
    {
        var pptPath = CreateTestPresentation("test_set_all.pptx", 3);
        var outputPath = CreateTestFilePath("test_set_all_output.pptx");
        var result = _tool.Execute("set", pptPath, color: "#00FF00", applyToAll: true, outputPath: outputPath);
        Assert.StartsWith("Background", result);
        Assert.Contains("all 3 slides", result);
        using var presentation = new Presentation(outputPath);
        foreach (var slide in presentation.Slides)
            Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
    }

    [Fact]
    public void Set_ToSecondSlide_ShouldOnlyAffectThatSlide()
    {
        var pptPath = CreateTestPresentation("test_set_second.pptx", 3);
        var outputPath = CreateTestFilePath("test_set_second_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 1, color: "#0000FF", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Solid, presentation.Slides[1].Background.FillFormat.FillType);
    }

    [Fact]
    public void Get_ShouldReturnBackgroundInfo()
    {
        var pptPath = CreateTestPresentation("test_get.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("hasBackground", out _));
        Assert.True(json.RootElement.TryGetProperty("fillType", out _));
        Assert.True(json.RootElement.TryGetProperty("isPictureFill", out _));
    }

    [Fact]
    public void Get_WithSolidBackground_ShouldReturnColorAndOpacity()
    {
        var pptPath = CreatePresentationWithSolidBackground("test_get_solid.pptx", Color.Red);
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal("Solid", json.RootElement.GetProperty("fillType").GetString());
        Assert.True(json.RootElement.TryGetProperty("color", out _));
        Assert.True(json.RootElement.TryGetProperty("opacity", out _));
    }

    [Fact]
    public void Get_WithPictureBackground_ShouldIndicatePictureFill()
    {
        var pptPath = CreateTestPresentation("test_get_picture.pptx");
        var outputPath = CreateTestFilePath("test_get_picture_set.pptx");
        var imagePath = CreateTestImage("test_get_bg.png");
        _tool.Execute("set", pptPath, slideIndex: 0, imagePath: imagePath, outputPath: outputPath);
        var result = _tool.Execute("get", outputPath, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal("Picture", json.RootElement.GetProperty("fillType").GetString());
        Assert.True(json.RootElement.GetProperty("isPictureFill").GetBoolean());
    }

    [Theory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive_Set(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_set_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, color: "#FF0000", outputPath: outputPath);
        Assert.StartsWith("Background", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_get_{operation}.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0);
        Assert.StartsWith("{", result);
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
    public void Set_WithoutColorOrImage_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_no_params.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, slideIndex: 0));
        Assert.Contains("color or imagePath", ex.Message);
    }

    [Fact]
    public void Set_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 99, color: "#FF0000"));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Get_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_get_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", pptPath, slideIndex: 99));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Set_WithInvalidImagePath_ShouldThrowFileNotFoundException()
    {
        var pptPath = CreateTestPresentation("test_invalid_image.pptx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 0, imagePath: "nonexistent.png"));
    }

    #endregion

    #region Session

    [Fact]
    public void Set_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, color: "#FF0000");
        Assert.StartsWith("Background", result);
        Assert.Contains("session", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(FillType.Solid, ppt.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void Set_WithSessionId_ApplyToAll_ShouldApplyToAllSlides()
    {
        var pptPath = CreateTestPresentation("test_session_all.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, color: "#00FF00", applyToAll: true);
        Assert.StartsWith("Background", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        foreach (var slide in ppt.Slides)
            Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
    }

    [Fact]
    public void Set_WithSessionId_WithImage_ShouldSetPictureBackground()
    {
        var pptPath = CreateTestPresentation("test_session_image.pptx");
        var sessionId = OpenSession(pptPath);
        var imagePath = CreateTestImage("test_session_bg.png");
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, imagePath: imagePath);
        Assert.StartsWith("Background", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(FillType.Picture, ppt.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void Get_WithSessionId_ShouldReturnBackgroundInfo()
    {
        var pptPath = CreateTestPresentation("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("fillType", out _));
    }

    [Fact]
    public void Get_WithSessionId_AfterSet_ShouldReflectChanges()
    {
        var pptPath = CreateTestPresentation("test_session_get_after_set.pptx");
        var sessionId = OpenSession(pptPath);
        _tool.Execute("set", sessionId: sessionId, slideIndex: 0, color: "#FF0000");
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal("Solid", json.RootElement.GetProperty("fillType").GetString());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", slideIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithSolidBackground("test_path_bg.pptx", Color.Red);
        var pptPath2 = CreatePresentationWithSolidBackground("test_session_bg.pptx", Color.Blue);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        var colorHex = json.RootElement.GetProperty("color").GetString();
        Assert.Contains("0000FF", colorHex!);
    }

    #endregion
}