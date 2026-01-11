using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for clearing slide content in PowerPoint presentations.
/// </summary>
public class ClearPptSlideHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "clear";

    /// <summary>
    ///     Clears all content from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex (0-based)
    /// </param>
    /// <returns>Success message with clear details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        while (slide.Shapes.Count > 0)
            slide.Shapes.RemoveAt(0);

        MarkModified(context);

        return Success($"Cleared all shapes from slide {slideIndex}.");
    }
}
