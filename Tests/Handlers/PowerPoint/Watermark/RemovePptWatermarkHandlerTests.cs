using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Watermark;

public class RemovePptWatermarkHandlerTests : PptHandlerTestBase
{
    private readonly RemovePptWatermarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Remove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region Remove Watermarks

    [Fact]
    public void Execute_WithWatermarks_ReturnsSuccessResult()
    {
        var pres = CreateEmptyPresentation();
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        shape.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_test";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("1 watermark", result.Message);
    }

    [Fact]
    public void Execute_WithWatermarks_MarksContextModified()
    {
        var pres = CreateEmptyPresentation();
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        shape.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_test";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_NoWatermarks_ReturnsZeroCount()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("0 watermark", result.Message);
    }

    [Fact]
    public void Execute_NoWatermarks_DoesNotMarkModified()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ShouldNotRemoveNonWatermarkShapes()
    {
        var pres = CreateEmptyPresentation();
        var normal = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        normal.Name = "NormalShape";
        var wm = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        wm.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_test";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Single(pres.Slides[0].Shapes);
    }

    [Fact]
    public void Execute_RemovesMultipleWatermarks()
    {
        var pres = CreateEmptyPresentation();
        var wm1 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        wm1.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_one";
        var wm2 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        wm2.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}IMAGE_two";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("2 watermark", result.Message);
        Assert.Empty(pres.Slides[0].Shapes);
    }

    [Fact]
    public void Execute_RemovesWatermarksAcrossSlides()
    {
        var pres = CreatePresentationWithSlides(2);
        var wm1 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        wm1.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_slide0";
        var wm2 = pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        wm2.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_slide1";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("2 watermark", result.Message);
    }

    #endregion
}
