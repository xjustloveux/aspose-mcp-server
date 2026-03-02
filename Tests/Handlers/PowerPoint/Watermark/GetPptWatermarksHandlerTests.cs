using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Watermark;
using AsposeMcpServer.Results.PowerPoint.Watermark;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Watermark;

[SupportedOSPlatform("windows")]
public class GetPptWatermarksHandlerTests : PptHandlerTestBase
{
    private readonly GetPptWatermarksHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Get()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Get Watermarks

    [SkippableFact]
    public void Execute_NoWatermarks_ReturnsEmptyResult()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWatermarksPptResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Empty(result.Items);
    }

    [SkippableFact]
    public void Execute_WithTextWatermark_ReturnsWatermarkInfo()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        shape.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_test";
        shape.TextFrame.Text = "DRAFT";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWatermarksPptResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Equal("text", result.Items[0].Type);
        Assert.Equal("DRAFT", result.Items[0].Text);
    }

    [SkippableFact]
    public void Execute_WithImageWatermark_ReturnsImageType()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        shape.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}IMAGE_test";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWatermarksPptResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Equal("image", result.Items[0].Type);
    }

    [SkippableFact]
    public void Execute_DoesNotMarkModified()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsMessage()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWatermarksPptResult>(res);
        Assert.NotNull(result.Message);
        Assert.Contains("watermark", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectCount()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var wm1 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        wm1.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_one";
        var wm2 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        wm2.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}IMAGE_two";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWatermarksPptResult>(res);
        Assert.Equal(result.Items.Count, result.Count);
        Assert.Equal(2, result.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        var wm = pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        wm.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_slide1";
        wm.TextFrame.Text = "WM";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWatermarksPptResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Equal(1, result.Items[0].SlideIndex);
    }

    [SkippableFact]
    public void Execute_IgnoresNonWatermarkShapes()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var normal = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        normal.Name = "NormalShape";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWatermarksPptResult>(res);
        Assert.Equal(0, result.Count);
    }

    #endregion
}
