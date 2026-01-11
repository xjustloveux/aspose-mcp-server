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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var layoutStr = parameters.GetRequired<string>("layout");

        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

        var layout = PptLayoutHelper.FindLayoutByType(presentation, layoutStr);
        presentation.Slides[slideIndex].LayoutSlide = layout;

        MarkModified(context);

        return Success($"Layout '{layoutStr}' set for slide {slideIndex}.");
    }
}
