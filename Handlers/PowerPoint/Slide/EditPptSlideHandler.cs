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
        var p = ExtractEditPptSlideParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        if (p.LayoutIndex.HasValue)
        {
            if (p.LayoutIndex.Value < 0 || p.LayoutIndex.Value >= presentation.LayoutSlides.Count)
                throw new ArgumentException(
                    $"layoutIndex must be between 0 and {presentation.LayoutSlides.Count - 1}");
            slide.LayoutSlide = presentation.LayoutSlides[p.LayoutIndex.Value];
        }

        MarkModified(context);

        return Success($"Slide {p.SlideIndex} updated.");
    }

    /// <summary>
    ///     Extracts parameters for edit slide operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static EditPptSlideParameters ExtractEditPptSlideParameters(OperationParameters parameters)
    {
        return new EditPptSlideParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetOptional<int?>("layoutIndex"));
    }

    /// <summary>
    ///     Parameters for edit slide operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="LayoutIndex">The layout index to apply.</param>
    private sealed record EditPptSlideParameters(int SlideIndex, int? LayoutIndex);
}
