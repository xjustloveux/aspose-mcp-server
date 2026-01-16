using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for setting layout on a PowerPoint slide.
/// </summary>
public class SetLayoutHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set";

    /// <summary>
    ///     Sets layout for a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, layout
    /// </param>
    /// <returns>Success message with layout details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSetLayoutParameters(parameters);

        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(p.SlideIndex, presentation.Slides.Count, "slide");

        var layout = PptLayoutHelper.FindLayoutByType(presentation, p.Layout);
        presentation.Slides[p.SlideIndex].LayoutSlide = layout;

        MarkModified(context);

        return Success($"Layout '{p.Layout}' set for slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts set layout parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set layout parameters.</returns>
    private static SetLayoutParameters ExtractSetLayoutParameters(OperationParameters parameters)
    {
        return new SetLayoutParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<string>("layout")
        );
    }

    /// <summary>
    ///     Record for holding set layout parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="Layout">The layout type string.</param>
    private record SetLayoutParameters(int SlideIndex, string Layout);
}
