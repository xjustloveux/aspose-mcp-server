using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Notes;

public class ClearNotesHandlerTests : PptHandlerTestBase
{
    private readonly ClearNotesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Clear()
    {
        Assert.Equal("clear", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithNotes(int slideCount)
    {
        var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++) presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

        foreach (var slide in presentation.Slides)
        {
            var notesSlide = slide.NotesSlideManager.AddNotesSlide();
            notesSlide.NotesTextFrame.Text = "Test notes";
        }

        return presentation;
    }

    #endregion

    #region Basic Clear Notes Operations

    [Fact]
    public void Execute_ClearsAllNotes()
    {
        var presentation = CreatePresentationWithNotes(3);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("cleared", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_ClearsSpecificSlideNotes()
    {
        var presentation = CreatePresentationWithNotes(3);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", new[] { 0, 2 } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2 targeted", result);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var presentation = CreatePresentationWithSlides(2);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", new[] { 999 } }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoNotes_ReportsZeroCleared()
    {
        var presentation = CreatePresentationWithSlides(2);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0 slides", result);
    }

    #endregion
}
