using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Hyperlink;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Hyperlink;

public class GetPptHyperlinksHandlerTests : PptHandlerTestBase
{
    private readonly GetPptHyperlinksHandler _handler = new();

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

    #region Helper Methods

    private static Presentation CreatePresentationWithHyperlink()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);
        shape.HyperlinkClick = new Aspose.Slides.Hyperlink("https://example.com");
        return pres;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsHyperlinks()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("totalCount", result);
        Assert.Contains("slides", result.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithSlideIndex_ReturnsHyperlinksForSlide()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"slideIndex\": 0", result);
        Assert.Contains("hyperlinks", result.ToLower());
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
