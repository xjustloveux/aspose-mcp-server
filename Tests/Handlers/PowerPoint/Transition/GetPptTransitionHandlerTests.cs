using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Handlers.PowerPoint.Transition;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Transition;

public class GetPptTransitionHandlerTests : PptHandlerTestBase
{
    private readonly GetPptTransitionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Multiple Slides

    [Fact]
    public void Execute_WithSlideIndex_ReturnsCorrectSlideInfo()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].SlideShowTransition.Type = TransitionType.Wipe;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.Equal("Wipe", json.RootElement.GetProperty("type").GetString());
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

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithTransition(TransitionType type)
    {
        var pres = new Presentation();
        pres.Slides[0].SlideShowTransition.Type = type;
        return pres;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsTransitionInfo()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("type", out _));
        Assert.True(json.RootElement.TryGetProperty("hasTransition", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithNoTransition_ReturnsNone()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal("None", json.RootElement.GetProperty("type").GetString());
        Assert.False(json.RootElement.GetProperty("hasTransition").GetBoolean());
    }

    [Fact]
    public void Execute_WithTransition_ReturnsTransitionType()
    {
        var pres = CreatePresentationWithTransition(TransitionType.Fade);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal("Fade", json.RootElement.GetProperty("type").GetString());
        Assert.True(json.RootElement.GetProperty("hasTransition").GetBoolean());
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsAdvanceSettings()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("advanceOnClick", out _));
        Assert.True(json.RootElement.TryGetProperty("advanceAfter", out _));
        Assert.True(json.RootElement.TryGetProperty("advanceAfterSeconds", out _));
    }

    #endregion
}
