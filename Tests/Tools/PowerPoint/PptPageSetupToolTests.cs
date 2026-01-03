using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptPageSetupToolTests : TestBase
{
    private readonly PptPageSetupTool _tool;

    public PptPageSetupToolTests()
    {
        _tool = new PptPageSetupTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void SetSlideSize_WithPreset_ShouldSetSlideSize()
    {
        var pptPath = CreateTestPresentation("test_set_slide_size.pptx");
        var outputPath = CreateTestFilePath("test_set_slide_size_output.pptx");
        var result = _tool.Execute("set_size", pptPath, preset: "OnScreen16x9", outputPath: outputPath);
        Assert.Contains("Slide size set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideSizeType.OnScreen, presentation.SlideSize.Type);
    }

    [Fact]
    public void SetSlideSize_WithCustomSize_ShouldSetCustomDimensions()
    {
        var pptPath = CreateTestPresentation("test_set_custom_size.pptx");
        var outputPath = CreateTestFilePath("test_set_custom_size_output.pptx");
        var result = _tool.Execute("set_size", pptPath, preset: "Custom", width: 720, height: 540,
            outputPath: outputPath);
        Assert.Contains("Custom", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(720, presentation.SlideSize.Size.Width, 1);
        Assert.Equal(540, presentation.SlideSize.Size.Height, 1);
    }

    [Fact]
    public void SetSlideSize_CustomWithoutDimensions_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_custom_no_dims.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("set_size", pptPath, preset: "Custom"));
    }

    [Fact]
    public void SetSlideSize_WithOutOfRangeWidth_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_size_out_of_range.pptx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_size", pptPath, preset: "Custom", width: 10000, height: 540));
    }

    [Fact]
    public void SetSlideOrientation_Portrait_ShouldSwapWidthAndHeight()
    {
        var pptPath = CreateTestPresentation("test_set_orientation.pptx");
        var outputPath = CreateTestFilePath("test_set_orientation_output.pptx");
        var result = _tool.Execute("set_orientation", pptPath, orientation: "Portrait", outputPath: outputPath);
        Assert.Contains("Portrait", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.SlideSize.Size.Height > presentation.SlideSize.Size.Width);
    }

    [Fact]
    public void SetSlideOrientation_Landscape_ShouldKeepOrSwapToLandscape()
    {
        // Arrange - Create a portrait presentation first
        var pptPath = CreateTestFilePath("test_landscape_orientation.pptx");
        using (var ppt = new Presentation())
        {
            ppt.SlideSize.SetSize(540, 720, SlideSizeScaleType.DoNotScale);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_landscape_orientation_output.pptx");
        var result = _tool.Execute("set_orientation", pptPath, orientation: "Landscape", outputPath: outputPath);
        Assert.Contains("Landscape", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.SlideSize.Size.Width > presentation.SlideSize.Size.Height);
    }

    [Fact]
    public void SetFooter_ShouldSetFooterText()
    {
        var pptPath = CreateTestPresentation("test_set_footer.pptx");
        var outputPath = CreateTestFilePath("test_set_footer_output.pptx");
        var result = _tool.Execute("set_footer", pptPath, footerText: "Footer Text", outputPath: outputPath);
        Assert.Contains("slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetFooter_WithDateText_ShouldSetDateText()
    {
        var pptPath = CreateTestPresentation("test_footer_date.pptx");
        var outputPath = CreateTestFilePath("test_footer_date_output.pptx");
        var result = _tool.Execute("set_footer", pptPath, footerText: "Footer", dateText: "2024-12-28",
            outputPath: outputPath);
        Assert.Contains("slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetFooter_WithSlideIndices_ShouldSetForSpecificSlides()
    {
        var pptPath = CreateTestFilePath("test_footer_indices.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_footer_indices_output.pptx");
        var result = _tool.Execute("set_footer", pptPath, footerText: "Footer for specific slides",
            slideIndices: [0, 1], outputPath: outputPath);
        Assert.Contains("2 slide(s)", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetFooter_InvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_footer_invalid.pptx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_footer", pptPath, footerText: "Test", slideIndices: [99]));
    }

    [Fact]
    public void SetSlideNumbering_ShouldSetSlideNumbering()
    {
        var pptPath = CreateTestPresentation("test_slide_numbering.pptx");
        var outputPath = CreateTestFilePath("test_slide_numbering_output.pptx");
        _tool.Execute("set_slide_numbering", pptPath, showSlideNumber: true, outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public void SetSlideNumbering_WithFirstNumber_ShouldSetStartingNumber()
    {
        var pptPath = CreateTestPresentation("test_slide_numbering_first.pptx");
        var outputPath = CreateTestFilePath("test_slide_numbering_first_output.pptx");
        var result = _tool.Execute("set_slide_numbering", pptPath, showSlideNumber: true, firstNumber: 5,
            outputPath: outputPath);
        Assert.Contains("starting from 5", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(5, presentation.FirstSlideNumber);
    }

    [Fact]
    public void SetSlideNumbering_Hide_ShouldHideNumbers()
    {
        var pptPath = CreateTestPresentation("test_hide_numbering.pptx");
        var outputPath = CreateTestFilePath("test_hide_numbering_output.pptx");
        var result = _tool.Execute("set_slide_numbering", pptPath, showSlideNumber: false, outputPath: outputPath);
        Assert.Contains("hidden", result);
        Assert.True(File.Exists(outputPath));
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
    public void SetOrientation_MissingOrientation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_missing_orientation.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set_orientation", pptPath));
        Assert.Contains("orientation", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void SetSlideSize_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set_size.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_size", sessionId: sessionId, preset: "OnScreen16x9");
        Assert.Contains("Slide size set to", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(SlideSizeType.OnScreen, ppt.SlideSize.Type);
    }

    [Fact]
    public void SetOrientation_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_orientation.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set_orientation", sessionId: sessionId, orientation: "Portrait");
        Assert.Contains("Portrait", result);
        Assert.True(ppt.SlideSize.Size.Height > ppt.SlideSize.Size.Width);
    }

    [Fact]
    public void SetFooter_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_footer.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_footer", sessionId: sessionId, footerText: "Session Footer");
        Assert.Contains("slide(s)", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void SetSlideNumbering_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_numbering.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_slide_numbering", sessionId: sessionId, showSlideNumber: true, firstNumber: 10);
        Assert.Contains("starting from 10", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(10, ppt.FirstSlideNumber);
    }

    #endregion
}