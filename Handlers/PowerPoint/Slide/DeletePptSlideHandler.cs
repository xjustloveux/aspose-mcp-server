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
        var p = ExtractDeletePptSlideParameters(parameters);

        var presentation = context.Document;

        if (p.SlideIndex < 0 || p.SlideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        if (presentation.Slides.Count == 1)
            throw new InvalidOperationException(
                "Cannot delete the last slide. A presentation must have at least one slide.");

        presentation.Slides.RemoveAt(p.SlideIndex);

        MarkModified(context);

        return Success($"Slide {p.SlideIndex} deleted ({presentation.Slides.Count} remaining).");
    }

    /// <summary>
    ///     Extracts parameters for delete slide operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeletePptSlideParameters ExtractDeletePptSlideParameters(OperationParameters parameters)
    {
        return new DeletePptSlideParameters(parameters.GetRequired<int>("slideIndex"));
    }

    /// <summary>
    ///     Parameters for delete slide operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    private record DeletePptSlideParameters(int SlideIndex);
}
