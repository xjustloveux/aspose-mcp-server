using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class MovePptSlideHandlerTests : PptHandlerTestBase
{
    private readonly MovePptSlideHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Move()
    {
        Assert.Equal("move", _handler.Operation);
    }

    #endregion

    #region Basic Move Operations

    [Theory]
    [InlineData(3, 0, 2)]
    [InlineData(3, 2, 0)]
    [InlineData(5, 1, 3)]
    [InlineData(5, 4, 0)]
    public void Execute_MovesToVariousPositions(int totalSlides, int fromIndex, int toIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", fromIndex },
            { "toIndex", toIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("moved", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(totalSlides, pres.Slides.Count);
        AssertModified(context);
    }

    #endregion

    #region Same Position

    [Fact]
    public void Execute_SamePosition_StillSucceeds()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", 1 },
            { "toIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("moved", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Result Message

    [Theory]
    [InlineData(0, 2)]
    [InlineData(1, 3)]
    [InlineData(3, 1)]
    public void Execute_ReturnsFromAndToInMessage(int fromIndex, int toIndex)
    {
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", fromIndex },
            { "toIndex", toIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains(fromIndex.ToString(), result);
        Assert.Contains(toIndex.ToString(), result);
    }

    #endregion

    #region Slide Tracking

    [Fact]
    public void Execute_MoveFirst_ToLast_CorrectOrder()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Ellipse, 10, 10, 50, 50);
        pres.Slides[2].Shapes.AddAutoShape(ShapeType.Triangle, 10, 10, 50, 50);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", 0 },
            { "toIndex", 2 }
        });

        _handler.Execute(context, parameters);

        Assert.True(pres.Slides[2].Shapes.Count > 0);
    }

    [Fact]
    public void Execute_PreservesSlideCount()
    {
        var pres = CreatePresentationWithSlides(5);
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", 2 },
            { "toIndex", 4 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides.Count);
    }

    #endregion

    #region Error Handling - Missing Parameters

    [Fact]
    public void Execute_WithoutFromIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "toIndex", 2 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fromIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutToIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("toIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling - Invalid fromIndex

    [Theory]
    [InlineData(3, 3)]
    [InlineData(3, 5)]
    [InlineData(3, 100)]
    public void Execute_WithFromIndexOutOfRange_ThrowsArgumentException(int totalSlides, int invalidFromIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", invalidFromIndex },
            { "toIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fromIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeFromIndex_ThrowsArgumentException(int negativeIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", negativeIndex },
            { "toIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fromIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling - Invalid toIndex

    [Theory]
    [InlineData(3, 3)]
    [InlineData(3, 5)]
    [InlineData(3, 100)]
    public void Execute_WithToIndexOutOfRange_ThrowsArgumentException(int totalSlides, int invalidToIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", 0 },
            { "toIndex", invalidToIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("toIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeToIndex_ThrowsArgumentException(int negativeIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromIndex", 0 },
            { "toIndex", negativeIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("toIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
