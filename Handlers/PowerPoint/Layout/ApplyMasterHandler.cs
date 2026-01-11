using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for applying a master slide layout to specified slides.
/// </summary>
public class ApplyMasterHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "apply_master";

    /// <summary>
    ///     Applies a master slide layout to specified slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: masterIndex, layoutIndex
    ///     Optional: slideIndices (JSON array)
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var masterIndex = parameters.GetRequired<int>("masterIndex");
        var layoutIndex = parameters.GetRequired<int>("layoutIndex");
        var slideIndicesJson = parameters.GetOptional<string?>("slideIndices");

        var presentation = context.Document;

        PowerPointHelper.ValidateCollectionIndex(masterIndex, presentation.Masters.Count, "master");
        var master = presentation.Masters[masterIndex];
        PowerPointHelper.ValidateCollectionIndex(layoutIndex, master.LayoutSlides.Count, "layout");

        var slideIndicesArray = PptLayoutHelper.ParseSlideIndicesJson(slideIndicesJson);
        var targets = slideIndicesArray?.Length > 0
            ? slideIndicesArray
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        PptLayoutHelper.ValidateSlideIndices(targets, presentation.Slides.Count);

        var layout = master.LayoutSlides[layoutIndex];
        foreach (var idx in targets)
            presentation.Slides[idx].LayoutSlide = layout;

        MarkModified(context);

        return Success($"Master {masterIndex} / Layout {layoutIndex} applied to {targets.Length} slides.");
    }
}
