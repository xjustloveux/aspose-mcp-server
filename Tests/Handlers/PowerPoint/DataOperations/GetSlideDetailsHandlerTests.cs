using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Results.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.DataOperations;

public class GetSlideDetailsHandlerTests : PptHandlerTestBase
{
    private readonly GetSlideDetailsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetSlideDetails()
    {
        Assert.Equal("get_slide_details", _handler.Operation);
    }

    #endregion

    #region Basic Get Slide Details Operations

    [Fact]
    public void Execute_ReturnsSlideDetails()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.Equal(0, result.SlideIndex);
        Assert.True(result.ShapesCount >= 0);
    }

    [Fact]
    public void Execute_ReturnsLayoutInfo()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.NotNull(result);
    }

    [Fact]
    public void Execute_ReturnsTransitionInfo()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.NotNull(result);
    }

    [Fact]
    public void Execute_ReturnsAnimationsCount()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.True(result.AnimationsCount >= 0);
    }

    [Fact]
    public void Execute_WithInvalidIndex_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
