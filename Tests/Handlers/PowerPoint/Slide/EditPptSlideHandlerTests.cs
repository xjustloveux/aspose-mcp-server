using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class EditPptSlideHandlerTests : PptHandlerTestBase
{
    private readonly EditPptSlideHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSlideIndexInMessage()
    {
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3", result.Message);
    }

    #endregion

    #region Error Handling - Missing Parameter

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_WithSlideIndex_ReturnsSuccess()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_EditsSlideAtVariousIndices(int slideIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", slideIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Layout Change

    [Fact]
    public void Execute_WithLayoutIndex_ChangesLayout()
    {
        var pres = CreatePresentationWithSlides(1);
        var originalLayoutName = pres.Slides[0].LayoutSlide.Name;
        var layoutCount = pres.LayoutSlides.Count;
        if (layoutCount < 2) return;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "layoutIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.NotEqual(originalLayoutName, pres.Slides[0].LayoutSlide.Name);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_WithVariousLayoutIndices_AppliesLayout(int layoutIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        if (pres.LayoutSlides.Count <= layoutIndex) return;
        var targetLayoutName = pres.LayoutSlides[layoutIndex].Name;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "layoutIndex", layoutIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(targetLayoutName, pres.Slides[0].LayoutSlide.Name);
    }

    #endregion

    #region Content Preservation

    [Fact]
    public void Execute_PreservesSlideShapes()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var shapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(shapeCount, pres.Slides[0].Shapes.Count);
    }

    [Fact]
    public void Execute_PreservesOtherSlides()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        pres.Slides[2].Shapes.AddAutoShape(ShapeType.Ellipse, 10, 10, 50, 50);
        var slide0ShapeCount = pres.Slides[0].Shapes.Count;
        var slide2ShapeCount = pres.Slides[2].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(slide0ShapeCount, pres.Slides[0].Shapes.Count);
        Assert.Equal(slide2ShapeCount, pres.Slides[2].Shapes.Count);
    }

    #endregion

    #region Error Handling - Invalid slideIndex

    [Theory]
    [InlineData(3, 3)]
    [InlineData(3, 5)]
    [InlineData(3, 100)]
    public void Execute_WithSlideIndexOutOfRange_ThrowsException(int totalSlides, int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", invalidIndex }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeSlideIndex_ThrowsException(int negativeIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", negativeIndex }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Error Handling - Invalid layoutIndex

    [Fact]
    public void Execute_WithLayoutIndexOutOfRange_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        var layoutCount = pres.LayoutSlides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "layoutIndex", layoutCount + 10 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("layoutIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeLayoutIndex_ThrowsArgumentException(int negativeLayoutIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "layoutIndex", negativeLayoutIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("layoutIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
