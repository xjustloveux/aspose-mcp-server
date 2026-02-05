using AsposeMcpServer.Handlers.PowerPoint.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Watermark;

public class AddTextPptWatermarkHandlerTests : PptHandlerTestBase
{
    private readonly AddTextPptWatermarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddText()
    {
        Assert.Equal("add_text", _handler.Operation);
    }

    #endregion

    #region Add Text Watermark

    [Fact]
    public void Execute_AddsWatermarkToAllSlides_ReturnsSuccessResult()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DRAFT" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("3 slide(s)", result.Message);
    }

    [Fact]
    public void Execute_AddsWatermarkToAllSlides_MarksContextModified()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DRAFT" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_CreatesNamedShape()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CONFIDENTIAL" }
        });

        _handler.Execute(context, parameters);

        var shapes = pres.Slides[0].Shapes;
        var hasWatermark = false;
        foreach (var shape in shapes)
            if (shape.Name != null && shape.Name.StartsWith(AddTextPptWatermarkHandler.WatermarkPrefix))
            {
                hasWatermark = true;
                break;
            }

        Assert.True(hasWatermark);
    }

    [Fact]
    public void Execute_WithMissingText_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithCustomFontSize_ReturnsSuccessResult()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "TEST" },
            { "fontSize", 72f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    [Fact]
    public void Execute_WithRotation_ReturnsSuccessResult()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "ROTATED" },
            { "rotation", -30f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    [Fact]
    public void Execute_WithCustomOpacity_ReturnsSuccessResult()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "SEMI-VISIBLE" },
            { "opacity", 64 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    [Fact]
    public void Execute_SingleSlide_ReturnsCorrectCount()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "SINGLE" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("1 slide(s)", result.Message);
    }

    [Fact]
    public void Execute_WatermarkShapeNameContainsTextPrefix()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CHECK" }
        });

        _handler.Execute(context, parameters);

        var shapes = pres.Slides[0].Shapes;
        var hasTextWatermark = false;
        foreach (var shape in shapes)
            if (shape.Name != null && shape.Name.Contains("_TEXT_"))
            {
                hasTextWatermark = true;
                break;
            }

        Assert.True(hasTextWatermark);
    }

    #endregion
}
