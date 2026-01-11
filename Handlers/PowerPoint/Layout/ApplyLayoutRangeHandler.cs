using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for applying layout to a range of slides.
/// </summary>
public class ApplyLayoutRangeHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "apply_layout_range";

    /// <summary>
    ///     Applies layout to a range of slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndices (JSON array), layout
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndicesJson = parameters.GetRequired<string>("slideIndices");
        var layoutStr = parameters.GetRequired<string>("layout");

        var slideIndicesArray = PptLayoutHelper.ParseSlideIndicesJson(slideIndicesJson);
        if (slideIndicesArray == null || slideIndicesArray.Length == 0)
            throw new ArgumentException("slideIndices is required for apply_layout_range operation");

        var presentation = context.Document;

        PptLayoutHelper.ValidateSlideIndices(slideIndicesArray, presentation.Slides.Count);

        var layout = PptLayoutHelper.FindLayoutByType(presentation, layoutStr);

        foreach (var idx in slideIndicesArray)
            presentation.Slides[idx].LayoutSlide = layout;

        MarkModified(context);

        return Success($"Layout '{layoutStr}' applied to {slideIndicesArray.Length} slide(s).");
    }
}
