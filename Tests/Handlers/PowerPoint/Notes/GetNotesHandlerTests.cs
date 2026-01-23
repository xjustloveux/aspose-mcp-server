using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Results.PowerPoint.Notes;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetNotesResult>(res);

        Assert.Equal(3, result.Count);
        Assert.NotNull(result.Slides);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetNotesResult>(res);

        Assert.Equal(0, result.SlideIndex);
        Assert.NotNull(result.HasNotes);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetNotesResult>(res);

        Assert.NotNull(result.HasNotes);
        Assert.IsType<bool>(result.HasNotes);
    }

    #endregion
}
