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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var notes = parameters.GetRequired<string>("notes");

        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

        var slide = presentation.Slides[slideIndex];
        var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
        notesSlide.NotesTextFrame.Text = notes;

        MarkModified(context);

        return Success($"Notes set for slide {slideIndex}.");
    }
}
