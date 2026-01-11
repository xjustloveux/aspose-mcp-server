using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Notes;

public class GetNotesHandlerTests : PptHandlerTestBase
{
    private readonly GetNotesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Get Notes Operations

    [Fact]
    public void Execute_ReturnsAllSlidesNotes()
    {
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 3", result);
        Assert.Contains("slides", result);
    }

    [Fact]
    public void Execute_WithSlideIndex_ReturnsSpecificSlideNotes()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"slideIndex\": 0", result);
        Assert.Contains("hasNotes", result);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_ReturnsHasNotesField()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("hasNotes", result);
    }

    #endregion
}
