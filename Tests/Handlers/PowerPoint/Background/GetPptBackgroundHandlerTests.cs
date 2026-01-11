using AsposeMcpServer.Handlers.PowerPoint.Background;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Background;

public class GetPptBackgroundHandlerTests : PptHandlerTestBase
{
    private readonly GetPptBackgroundHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsBackgroundInfo()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slideIndex", result);
        Assert.Contains("hasBackground", result);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithSlideIndex_ReturnsSpecificSlideBackground()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"slideIndex\": 1", result);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    #endregion
}
