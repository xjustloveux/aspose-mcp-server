using AsposeMcpServer.Handlers.PowerPoint.Notes;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Notes;

public class SetNotesHandlerTests : PptHandlerTestBase
{
    private readonly SetNotesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Set()
    {
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Basic Set Notes Operations

    [Fact]
    public void Execute_SetsNotesForSlide()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "notes", "Test speaker notes" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("notes set", result.ToLower());
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_ReplacesExistingNotes()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var presentation = CreateEmptyPresentation();
        var slide = presentation.Slides[0];
        var notesSlide = slide.NotesSlideManager.AddNotesSlide();
        notesSlide.NotesTextFrame.Text = "Old notes";

        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "notes", "New notes" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("New notes", slide.NotesSlideManager.NotesSlide.NotesTextFrame.Text);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 999 },
            { "notes", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
