using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.PowerPoint.Shape;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

[SupportedOSPlatform("windows")]
public class GetPptShapeDetailsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptShapeDetailsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetShapeDetails()
    {
        SkipIfNotWindows();
        Assert.Equal("get_shape_details", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [SkippableFact]
    public void Execute_DoesNotModifyPresentation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides[0].Shapes.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_ReturnsShapeDetails()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.True(result.Index >= 0);
        Assert.NotNull(result.Type);
        Assert.True(result.X >= 0);
        Assert.True(result.Y >= 0);
        Assert.True(result.Width >= 0);
        Assert.True(result.Height >= 0);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectPosition()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 150, 250, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(150f, result.X);
        Assert.Equal(250f, result.Y);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectDimensions()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 350, 450);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(350f, result.Width);
        Assert.Equal(450f, result.Height);
    }

    #endregion

    #region AutoShape Details

    [SkippableFact]
    public void Execute_ForAutoShape_ReturnsShapeType()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        var autoShape = Assert.IsType<AutoShapeDetails>(result.Details);
        Assert.Equal("Ellipse", autoShape.ShapeType);
    }

    [SkippableFact]
    public void Execute_ForAutoShape_ReturnsFillType()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.FillFormat.FillType = FillType.Solid;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        var autoShape = Assert.IsType<AutoShapeDetails>(result.Details);
        Assert.Equal("Solid", autoShape.FillType);
    }

    #endregion

    #region Rotation and Hidden

    [SkippableFact]
    public void Execute_ReturnsRotation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Rotation = 45;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(45f, result.Rotation);
    }

    [SkippableFact]
    public void Execute_ReturnsHiddenStatus()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Hidden = true;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.True(result.Hidden);
    }

    #endregion

    #region AlternativeText

    [SkippableFact]
    public void Execute_ReturnsAlternativeText()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.AlternativeText = "Test alt text";
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal("Test alt text", result.AlternativeText);
    }

    [SkippableFact]
    public void Execute_WithEmptyAlternativeText_ReturnsNull()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Null(result.AlternativeText);
    }

    #endregion

    #region Flip Properties

    [SkippableFact]
    public void Execute_WithFlipHorizontal_ReturnsTrue()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Frame = new ShapeFrame(shape.X, shape.Y, shape.Width, shape.Height, NullableBool.True,
            shape.Frame.FlipV, shape.Rotation);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.True(result.FlipHorizontal);
    }

    [SkippableFact]
    public void Execute_WithFlipVertical_ReturnsTrue()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Frame = new ShapeFrame(shape.X, shape.Y, shape.Width, shape.Height, shape.Frame.FlipH,
            NullableBool.True, shape.Rotation);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.True(result.FlipVertical);
    }

    [SkippableFact]
    public void Execute_WithNoFlip_ReturnsNullOrFalse()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.True(result.FlipHorizontal is null or false);
        Assert.True(result.FlipVertical is null or false);
    }

    #endregion

    #region Slide Index

    [SkippableFact]
    public void Execute_WithSlideIndex_GetsFromSpecificSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Triangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(0, result.Index);
    }

    [SkippableFact]
    public void Execute_DefaultSlideIndex_GetsFromFirstSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(0, result.Index);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
