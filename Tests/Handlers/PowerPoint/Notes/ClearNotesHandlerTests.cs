using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Notes;

public class ClearNotesHandlerTests : PptHandlerTestBase
{
    private static readonly int[] SlideIndicesZeroTwo = [0, 2];
    private static readonly int[] InvalidSlideIndex = [999];

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
            foreach (var slide in presentation.Slides)
            {
                var notesSlide = slide.NotesSlideManager.NotesSlide;
                Assert.True(notesSlide == null || string.IsNullOrEmpty(notesSlide.NotesTextFrame.Text),
                    $"Notes should be cleared on slide {presentation.Slides.IndexOf(slide)}");
            }

        AssertModified(context);
    }

    [Fact]
    public void Execute_ClearsSpecificSlideNotes()
    {
        var presentation = CreatePresentationWithNotes(3);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", SlideIndicesZeroTwo }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var notes0 = presentation.Slides[0].NotesSlideManager.NotesSlide;
            Assert.True(notes0 == null || string.IsNullOrEmpty(notes0.NotesTextFrame.Text),
                "Notes on slide 0 should be cleared");
            var notes2 = presentation.Slides[2].NotesSlideManager.NotesSlide;
            Assert.True(notes2 == null || string.IsNullOrEmpty(notes2.NotesTextFrame.Text),
                "Notes on slide 2 should be cleared");
            var notes1 = presentation.Slides[1].NotesSlideManager.NotesSlide;
            Assert.False(string.IsNullOrEmpty(notes1?.NotesTextFrame.Text),
                "Notes on slide 1 should remain untouched");
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var presentation = CreatePresentationWithSlides(2);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", InvalidSlideIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoNotes_SlidesRemainWithoutNotes()
    {
        var presentation = CreatePresentationWithSlides(2);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        foreach (var slide in presentation.Slides)
        {
            var notesSlide = slide.NotesSlideManager.NotesSlide;
            Assert.True(notesSlide == null || string.IsNullOrEmpty(notesSlide.NotesTextFrame.Text),
                "Slides without notes should remain without notes");
        }
    }

    #endregion
}
