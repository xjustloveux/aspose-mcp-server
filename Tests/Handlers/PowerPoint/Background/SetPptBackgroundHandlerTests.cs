using AsposeMcpServer.Handlers.PowerPoint.Background;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Background;

public class SetPptBackgroundHandlerTests : PptHandlerTestBase
{
    private readonly SetPptBackgroundHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Set()
    {
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsBackgroundColor()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("background", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSlideIndex_SetsBackgroundOnSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "color", "#00FF00" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 1", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithApplyToAll_SetsBackgroundOnAllSlides()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "#0000FF" },
            { "applyToAll", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("all", result.ToLower());
        Assert.Contains("3", result);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutColorOrImage_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "color", "#FF0000" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
