using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slideIndex", result);
        Assert.Contains("shapesCount", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("layout", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("transition", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("animationsCount", result);
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
