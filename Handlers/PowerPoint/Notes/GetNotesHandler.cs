using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Notes;

namespace AsposeMcpServer.Handlers.PowerPoint.Notes;

/// <summary>
///     Handler for getting notes from PowerPoint slides.
/// </summary>
[ResultType(typeof(GetNotesResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetNotesParameters(parameters);

        var presentation = context.Document;

        if (p.SlideIndex.HasValue)
        {
            PowerPointHelper.ValidateCollectionIndex(p.SlideIndex.Value, presentation.Slides.Count, "slide");

            var notesSlide = presentation.Slides[p.SlideIndex.Value].NotesSlideManager.NotesSlide;
            var notesText = notesSlide?.NotesTextFrame?.Text;

            return new GetNotesResult
            {
                SlideIndex = p.SlideIndex.Value,
                HasNotes = !string.IsNullOrWhiteSpace(notesText),
                Notes = notesText
            };
        }

        List<SlideNotesInfo> notesList = [];
        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var notesSlide = presentation.Slides[i].NotesSlideManager.NotesSlide;
            var notesText = notesSlide?.NotesTextFrame?.Text;

            notesList.Add(new SlideNotesInfo
            {
                SlideIndex = i,
                HasNotes = !string.IsNullOrWhiteSpace(notesText),
                Notes = notesText
            });
        }

        return new GetNotesResult
        {
            Count = presentation.Slides.Count,
            Slides = notesList
        };
    }

    /// <summary>
    ///     Extracts get notes parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get notes parameters.</returns>
    private static GetNotesParameters ExtractGetNotesParameters(OperationParameters parameters)
    {
        return new GetNotesParameters(parameters.GetOptional<int?>("slideIndex"));
    }

    /// <summary>
    ///     Record for holding get notes parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    private sealed record GetNotesParameters(int? SlideIndex);
}
