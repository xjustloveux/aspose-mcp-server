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

    private string CreateMultiSlidePresentation(string fileName, int slideCount = 3)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePortraitPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        ppt.SlideSize.SetSize(540, 720, SlideSizeScaleType.DoNotScale);
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void SetSize_WithPreset_ShouldSetSlideSize()
    {
        var pptPath = CreateTestPresentation("test_set_size.pptx");
        var outputPath = CreateTestFilePath("test_set_size_output.pptx");
        var result = _tool.Execute("set_size", pptPath, preset: "OnScreen16x9", outputPath: outputPath);
        Assert.StartsWith("Slide size set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideSizeType.OnScreen, presentation.SlideSize.Type);
    }

    [Fact]
    public void SetSize_WithCustomSize_ShouldSetCustomDimensions()
    {
        var pptPath = CreateTestPresentation("test_set_custom_size.pptx");
        var outputPath = CreateTestFilePath("test_set_custom_size_output.pptx");
        var result = _tool.Execute("set_size", pptPath, preset: "Custom", width: 720, height: 540,
            outputPath: outputPath);
        Assert.StartsWith("Slide size set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(720, presentation.SlideSize.Size.Width, 1);
        Assert.Equal(540, presentation.SlideSize.Size.Height, 1);
    }

    [Fact]
    public void SetOrientation_Portrait_ShouldSwapToPortrait()
    {
        var pptPath = CreateTestPresentation("test_set_portrait.pptx");
        var outputPath = CreateTestFilePath("test_set_portrait_output.pptx");
        var result = _tool.Execute("set_orientation", pptPath, orientation: "Portrait", outputPath: outputPath);
        Assert.StartsWith("Slide orientation set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.SlideSize.Size.Height > presentation.SlideSize.Size.Width);
    }

    [Fact]
    public void SetOrientation_Landscape_ShouldSwapToLandscape()
    {
        var pptPath = CreatePortraitPresentation("test_set_landscape.pptx");
        var outputPath = CreateTestFilePath("test_set_landscape_output.pptx");
        var result = _tool.Execute("set_orientation", pptPath, orientation: "Landscape", outputPath: outputPath);
        Assert.StartsWith("Slide orientation set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.SlideSize.Size.Width > presentation.SlideSize.Size.Height);
    }

    [Fact]
    public void SetFooter_ShouldSetFooterText()
    {
        var pptPath = CreateTestPresentation("test_set_footer.pptx");
        var outputPath = CreateTestFilePath("test_set_footer_output.pptx");
        var result = _tool.Execute("set_footer", pptPath, footerText: "Footer Text", outputPath: outputPath);
        Assert.StartsWith("Footer settings updated for", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetFooter_WithDateText_ShouldSetDateText()
    {
        var pptPath = CreateTestPresentation("test_footer_date.pptx");
        var outputPath = CreateTestFilePath("test_footer_date_output.pptx");
        var result = _tool.Execute("set_footer", pptPath, footerText: "Footer", dateText: "2024-12-28",
            outputPath: outputPath);
        Assert.StartsWith("Footer settings updated for", result);
    }

    [Fact]
    public void SetFooter_WithSlideIndices_ShouldSetForSpecificSlides()
    {
        var pptPath = CreateMultiSlidePresentation("test_footer_indices.pptx");
        var outputPath = CreateTestFilePath("test_footer_indices_output.pptx");
        var result = _tool.Execute("set_footer", pptPath, footerText: "Footer for specific slides",
            slideIndices: [0, 1], outputPath: outputPath);
        Assert.StartsWith("Footer settings updated for", result);
    }

    [Fact]
    public void SetSlideNumbering_Show_ShouldShowNumbers()
    {
        var pptPath = CreateTestPresentation("test_numbering_show.pptx");
        var outputPath = CreateTestFilePath("test_numbering_show_output.pptx");
        var result = _tool.Execute("set_slide_numbering", pptPath, showSlideNumber: true, outputPath: outputPath);
        Assert.StartsWith("Slide numbers", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetSlideNumbering_WithFirstNumber_ShouldSetStartingNumber()
    {
        var pptPath = CreateTestPresentation("test_numbering_first.pptx");
        var outputPath = CreateTestFilePath("test_numbering_first_output.pptx");
        var result = _tool.Execute("set_slide_numbering", pptPath, showSlideNumber: true, firstNumber: 5,
            outputPath: outputPath);
        Assert.StartsWith("Slide numbers", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(5, presentation.FirstSlideNumber);
    }

    [Fact]
    public void SetSlideNumbering_Hide_ShouldHideNumbers()
    {
        var pptPath = CreateTestPresentation("test_numbering_hide.pptx");
        var outputPath = CreateTestFilePath("test_numbering_hide_output.pptx");
        var result = _tool.Execute("set_slide_numbering", pptPath, showSlideNumber: false, outputPath: outputPath);
        Assert.StartsWith("Slide numbers", result);
    }

    [Theory]
    [InlineData("SET_SIZE")]
    [InlineData("Set_Size")]
    [InlineData("set_size")]
    public void Operation_ShouldBeCaseInsensitive_SetSize(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_size_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_size_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, preset: "OnScreen16x9", outputPath: outputPath);
        Assert.StartsWith("Slide size set to", result);
    }

    [Theory]
    [InlineData("SET_ORIENTATION")]
    [InlineData("Set_Orientation")]
    [InlineData("set_orientation")]
    public void Operation_ShouldBeCaseInsensitive_SetOrientation(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_orient_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_orient_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, orientation: "Portrait", outputPath: outputPath);
        Assert.StartsWith("Slide orientation set to", result);
    }

    [Theory]
    [InlineData("SET_FOOTER")]
    [InlineData("Set_Footer")]
    [InlineData("set_footer")]
    public void Operation_ShouldBeCaseInsensitive_SetFooter(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_footer_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_footer_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, footerText: "Footer", outputPath: outputPath);
        Assert.StartsWith("Footer settings updated for", result);
    }

    [Theory]
    [InlineData("SET_SLIDE_NUMBERING")]
    [InlineData("Set_Slide_Numbering")]
    [InlineData("set_slide_numbering")]
    public void Operation_ShouldBeCaseInsensitive_SetSlideNumbering(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_num_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_num_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, showSlideNumber: true, outputPath: outputPath);
        Assert.StartsWith("Slide numbers", result);
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
    public void SetSize_CustomWithoutDimensions_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_custom_no_dims.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set_size", pptPath, preset: "Custom"));
        Assert.Contains("width and height", ex.Message);
    }

    [Fact]
    public void SetSize_WithOutOfRangeWidth_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_size_out_of_range.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_size", pptPath, preset: "Custom", width: 10000, height: 540));
        Assert.Contains("Width must be between", ex.Message);
    }

    [Fact]
    public void SetSize_WithOutOfRangeHeight_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_height_out_of_range.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_size", pptPath, preset: "Custom", width: 540, height: 10000));
        Assert.Contains("Height must be between", ex.Message);
    }

    [Fact]
    public void SetOrientation_WithoutOrientation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_missing_orientation.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set_orientation", pptPath));
        Assert.Contains("orientation is required", ex.Message);
    }

    [Fact]
    public void SetFooter_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_footer_invalid.pptx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_footer", pptPath, footerText: "Test", slideIndices: [99]));
    }

    #endregion

    #region Session

    [Fact]
    public void SetSize_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_size.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_size", sessionId: sessionId, preset: "OnScreen16x9");
        Assert.StartsWith("Slide size set to", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(SlideSizeType.OnScreen, ppt.SlideSize.Type);
    }

    [Fact]
    public void SetOrientation_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_orientation.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_orientation", sessionId: sessionId, orientation: "Portrait");
        Assert.StartsWith("Slide orientation set to", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.True(ppt.SlideSize.Size.Height > ppt.SlideSize.Size.Width);
    }

    [Fact]
    public void SetFooter_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_footer.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_footer", sessionId: sessionId, footerText: "Session Footer");
        Assert.StartsWith("Footer settings updated for", result);
    }

    [Fact]
    public void SetSlideNumbering_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_numbering.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_slide_numbering", sessionId: sessionId, showSlideNumber: true, firstNumber: 10);
        Assert.StartsWith("Slide numbers", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(10, ppt.FirstSlideNumber);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("set_size", sessionId: "invalid_session", preset: "OnScreen16x9"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestPresentation("test_path_setup.pptx");
        var pptPath2 = CreateTestPresentation("test_session_setup.pptx");
        var sessionId = OpenSession(pptPath2);
        _tool.Execute("set_slide_numbering", pptPath1, sessionId, showSlideNumber: true, firstNumber: 99);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(99, ppt.FirstSlideNumber);
    }

    #endregion
}