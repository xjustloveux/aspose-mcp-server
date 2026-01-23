using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class DeletePptShapeHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptShapeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_DeletesShapeAtVariousIndices(int shapeIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 300, 100, 200, 100);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsIndexesInMessage()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("0", result.Message);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_DeletesFromSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var initialCount = pres.Slides[1].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, pres.Slides[1].Shapes.Count);
    }

    [Fact]
    public void Execute_DefaultSlideIndex_DeletesFromFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
