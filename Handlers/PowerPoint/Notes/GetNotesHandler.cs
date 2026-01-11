using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Notes;

/// <summary>
///     Handler for getting notes from PowerPoint slides.
/// </summary>
public class GetNotesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets notes from slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex (if not provided, returns all slides' notes)
    /// </param>
    /// <returns>JSON string containing notes information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional<int?>("slideIndex");

        var presentation = context.Document;

        if (slideIndex.HasValue)
        {
            PowerPointHelper.ValidateCollectionIndex(slideIndex.Value, presentation.Slides.Count, "slide");

            var notesSlide = presentation.Slides[slideIndex.Value].NotesSlideManager.NotesSlide;
            var notesText = notesSlide?.NotesTextFrame?.Text;

            return JsonResult(new
            {
                slideIndex = slideIndex.Value,
                hasNotes = !string.IsNullOrWhiteSpace(notesText),
                notes = notesText
            });
        }

        List<object> notesList = [];
        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var notesSlide = presentation.Slides[i].NotesSlideManager.NotesSlide;
            var notesText = notesSlide?.NotesTextFrame?.Text;

            notesList.Add(new
            {
                slideIndex = i,
                hasNotes = !string.IsNullOrWhiteSpace(notesText),
                notes = notesText
            });
        }

        return JsonResult(new
        {
            count = presentation.Slides.Count,
            slides = notesList
        });
    }
}
