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
        var p = ExtractApplyLayoutRangeParameters(parameters);

        var slideIndicesArray = PptLayoutHelper.ParseSlideIndicesJson(p.SlideIndicesJson);
        if (slideIndicesArray == null || slideIndicesArray.Length == 0)
            throw new ArgumentException("slideIndices is required for apply_layout_range operation");

        var presentation = context.Document;

        PptLayoutHelper.ValidateSlideIndices(slideIndicesArray, presentation.Slides.Count);

        var layout = PptLayoutHelper.FindLayoutByType(presentation, p.Layout);

        foreach (var idx in slideIndicesArray)
            presentation.Slides[idx].LayoutSlide = layout;

        MarkModified(context);

        return Success($"Layout '{p.Layout}' applied to {slideIndicesArray.Length} slide(s).");
    }

    /// <summary>
    ///     Extracts apply layout range parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted apply layout range parameters.</returns>
    private static ApplyLayoutRangeParameters ExtractApplyLayoutRangeParameters(OperationParameters parameters)
    {
        return new ApplyLayoutRangeParameters(
            parameters.GetRequired<string>("slideIndices"),
            parameters.GetRequired<string>("layout")
        );
    }

    /// <summary>
    ///     Record for holding apply layout range parameters.
    /// </summary>
    /// <param name="SlideIndicesJson">The JSON array of slide indices.</param>
    /// <param name="Layout">The layout type string.</param>
    private sealed record ApplyLayoutRangeParameters(string SlideIndicesJson, string Layout);
}
