using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

public class SetFooterHandlerTests : PptHandlerTestBase
{
    private static readonly int[] SlideIndicesZeroTwo = [0, 2];

    private readonly SetFooterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFooter()
    {
        Assert.Equal("set_footer", _handler.Operation);
    }

    #endregion

    #region Basic Set Footer Operations

    [Fact]
    public void Execute_SetsFooterText()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "My Footer" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("footer settings updated", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsDateText()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dateText", "2026-01-11" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("footer settings updated", result.ToLower());
    }

    [Fact]
    public void Execute_ShowsSlideNumber()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showSlideNumber", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("footer settings updated", result.ToLower());
    }

    [Fact]
    public void Execute_AppliesToAllSlides()
    {
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3 slide", result);
    }

    [Fact]
    public void Execute_AppliesToSpecificSlides()
    {
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" },
            { "slideIndices", SlideIndicesZeroTwo }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2 slide", result);
    }

    #endregion
}
