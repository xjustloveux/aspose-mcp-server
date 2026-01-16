using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Notes;

/// <summary>
///     Handler for setting notes on PowerPoint slides.
/// </summary>
public class SetNotesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set";

    /// <summary>
    ///     Sets (replaces) notes on a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, notes
    /// </param>
    /// <returns>Success message indicating notes were set.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSetNotesParameters(parameters);

        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(p.SlideIndex, presentation.Slides.Count, "slide");

        var slide = presentation.Slides[p.SlideIndex];
        var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
        notesSlide.NotesTextFrame.Text = p.Notes;

        MarkModified(context);

        return Success($"Notes set for slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts set notes parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set notes parameters.</returns>
    private static SetNotesParameters ExtractSetNotesParameters(OperationParameters parameters)
    {
        return new SetNotesParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<string>("notes"));
    }

    /// <summary>
    ///     Record for holding set notes parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="Notes">The notes content.</param>
    private sealed record SetNotesParameters(int SlideIndex, string Notes);
}
