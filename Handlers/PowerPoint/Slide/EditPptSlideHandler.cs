using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for editing slide properties in PowerPoint presentations.
/// </summary>
public class EditPptSlideHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits slide properties.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex (0-based)
    ///     Optional: layoutIndex (0-based, to change layout)
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var layoutIndex = parameters.GetOptional<int?>("layoutIndex");
        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        if (layoutIndex.HasValue)
        {
            if (layoutIndex.Value < 0 || layoutIndex.Value >= presentation.LayoutSlides.Count)
                throw new ArgumentException(
                    $"layoutIndex must be between 0 and {presentation.LayoutSlides.Count - 1}");
            slide.LayoutSlide = presentation.LayoutSlides[layoutIndex.Value];
        }

        MarkModified(context);

        return Success($"Slide {slideIndex} updated.");
    }
}
