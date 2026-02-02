using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Hyperlink;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Hyperlink;

public class AddPptHyperlinkHandlerTests : PptHandlerTestBase
{
    private readonly AddPptHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsHyperlinkWithUrl()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "url", "https://example.com" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.NotNull(shape.HyperlinkClick);
            Assert.Equal("https://example.com", shape.HyperlinkClick.ExternalUrl);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSlideTarget_AddsSlideHyperlink()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "slideTargetIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.NotNull(shape.HyperlinkClick);
            Assert.NotNull(shape.HyperlinkClick.TargetSlide);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithTextAndLinkText_AddsPartialHyperlink()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "url", "https://example.com" },
            { "text", "Click here to visit" },
            { "linkText", "here" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            var paragraph = shape.TextFrame.Paragraphs[0];
            Assert.True(paragraph.Portions.Count >= 2);
            var linkPortion = paragraph.Portions
                .First(p => p.Text == "here");
            Assert.NotNull(linkPortion.PortionFormat.HyperlinkClick);
            Assert.Equal("https://example.com", linkPortion.PortionFormat.HyperlinkClick.ExternalUrl);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "url", "https://example.com" }
        });

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
            { "url", "https://example.com" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithLinkTextNotInText_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "url", "https://example.com" },
            { "text", "Some text" },
            { "linkText", "notfound" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
