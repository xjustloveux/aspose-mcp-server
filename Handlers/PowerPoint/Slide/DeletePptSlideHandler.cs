using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for deleting slides from PowerPoint presentations.
/// </summary>
public class DeletePptSlideHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a slide from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex (0-based)
    /// </param>
    /// <returns>Success message with remaining slide count.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var presentation = context.Document;

        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        if (presentation.Slides.Count == 1)
            throw new InvalidOperationException(
                "Cannot delete the last slide. A presentation must have at least one slide.");

        presentation.Slides.RemoveAt(slideIndex);

        MarkModified(context);

        return Success($"Slide {slideIndex} deleted ({presentation.Slides.Count} remaining).");
    }
}
