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
        var p = ExtractApplyMasterParameters(parameters);

        var presentation = context.Document;

        PowerPointHelper.ValidateCollectionIndex(p.MasterIndex, presentation.Masters.Count, "master");
        var master = presentation.Masters[p.MasterIndex];
        PowerPointHelper.ValidateCollectionIndex(p.LayoutIndex, master.LayoutSlides.Count, "layout");

        var slideIndicesArray = PptLayoutHelper.ParseSlideIndicesJson(p.SlideIndicesJson);
        var targets = slideIndicesArray?.Length > 0
            ? slideIndicesArray
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        PptLayoutHelper.ValidateSlideIndices(targets, presentation.Slides.Count);

        var layout = master.LayoutSlides[p.LayoutIndex];
        foreach (var idx in targets)
            presentation.Slides[idx].LayoutSlide = layout;

        MarkModified(context);

        return Success($"Master {p.MasterIndex} / Layout {p.LayoutIndex} applied to {targets.Length} slides.");
    }

    /// <summary>
    ///     Extracts apply master parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted apply master parameters.</returns>
    private static ApplyMasterParameters ExtractApplyMasterParameters(OperationParameters parameters)
    {
        return new ApplyMasterParameters(
            parameters.GetRequired<int>("masterIndex"),
            parameters.GetRequired<int>("layoutIndex"),
            parameters.GetOptional<string?>("slideIndices")
        );
    }

    /// <summary>
    ///     Record for holding apply master parameters.
    /// </summary>
    /// <param name="MasterIndex">The master slide index.</param>
    /// <param name="LayoutIndex">The layout index.</param>
    /// <param name="SlideIndicesJson">The optional JSON array of slide indices.</param>
    private record ApplyMasterParameters(int MasterIndex, int LayoutIndex, string? SlideIndicesJson);
}
