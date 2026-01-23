using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class DuplicatePptSlideHandlerTests : PptHandlerTestBase
{
    private readonly DuplicatePptSlideHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Duplicate()
    {
        Assert.Equal("duplicate", _handler.Operation);
    }

    #endregion

    #region Basic Duplicate Operations

    [Theory]
    [InlineData(3, 0)]
    [InlineData(3, 1)]
    [InlineData(3, 2)]
    [InlineData(5, 4)]
    public void Execute_DuplicatesAtVariousIndices(int totalSlides, int slideIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", slideIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("duplicated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
        AssertModified(context);
    }

    #endregion

    #region Default Append Behavior

    [Fact]
    public void Execute_WithoutInsertAt_AppendsAtEnd()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(4, pres.Slides.Count);
        Assert.True(pres.Slides[3].Shapes.Count > 0);
    }

    #endregion

    #region InsertAt Position

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_WithVariousInsertAtPositions_InsertsCorrectly(int insertAt)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "insertAt", insertAt }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("duplicated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(4, pres.Slides.Count);
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

    #region Content Preservation

    [Fact]
    public void Execute_PreservesOriginalSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 200, 100);
        var originalShapeCount = pres.Slides[1].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(originalShapeCount, pres.Slides[1].Shapes.Count);
    }

    [SkippableFact]
    public void Execute_DuplicatedSlideHasSameShapeCount()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermark shapes");
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 50, 50);
        var originalShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(originalShapeCount, pres.Slides[1].Shapes.Count);
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

    [Fact]
    public void Execute_ReturnsTotalCountInMessage()
    {
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("6", result.Message);
    }

    #endregion

    #region Error Handling - Invalid slideIndex

    [Theory]
    [InlineData(3, 3)]
    [InlineData(3, 5)]
    [InlineData(3, 100)]
    public void Execute_WithSlideIndexOutOfRange_ThrowsArgumentException(int totalSlides, int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException(int negativeIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", negativeIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling - Invalid insertAt

    [Theory]
    [InlineData(3, 4)]
    [InlineData(3, 10)]
    public void Execute_WithInsertAtOutOfRange_ThrowsArgumentException(int totalSlides, int invalidInsertAt)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "insertAt", invalidInsertAt }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("insertAt", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeInsertAt_ThrowsArgumentException(int negativeInsertAt)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "insertAt", negativeInsertAt }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("insertAt", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
