using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Hyperlink;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Hyperlink;

public class EditPptHyperlinkHandlerTests : PptHandlerTestBase
{
    private readonly EditPptHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
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

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsHyperlinkUrl()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "url", "https://newurl.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ChangesToSlideTarget()
    {
        var pres = CreatePresentationWithHyperlink();
        pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "slideTargetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithRemoveHyperlink_RemovesHyperlink()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "removeHyperlink", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "url", "https://example.com" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "url", "https://example.com" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutUrlOrSlideTargetOrRemove_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideTargetIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithHyperlink();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "slideTargetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
