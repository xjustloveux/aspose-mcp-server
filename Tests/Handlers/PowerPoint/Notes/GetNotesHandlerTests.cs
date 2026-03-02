using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Results.PowerPoint.Notes;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Notes;

[SupportedOSPlatform("windows")]
public class GetNotesHandlerTests : PptHandlerTestBase
{
    private readonly GetNotesHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Get()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Get Notes Operations

    [SkippableFact]
    public void Execute_ReturnsAllSlidesNotes()
    {
        SkipIfNotWindows();
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetNotesResult>(res);

        Assert.Equal(3, result.Count);
        Assert.NotNull(result.Slides);
    }

    [SkippableFact]
    public void Execute_WithSlideIndex_ReturnsSpecificSlideNotes()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_ReturnsHasNotesField()
    {
        SkipIfNotWindows();
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
