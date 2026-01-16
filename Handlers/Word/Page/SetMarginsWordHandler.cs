using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for setting page margins in Word documents.
/// </summary>
public class SetMarginsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_margins";

    /// <summary>
    ///     Sets page margins for the specified sections.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: top, bottom, left, right (margins in points)
    ///     Optional: sectionIndex, sectionIndices
    /// </param>
    /// <returns>Success message with margin details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var setParams = ExtractSetMarginsParameters(parameters);

        var doc = context.Document;
        var sectionsToUpdate = WordPageHelper.GetTargetSections(doc, setParams.SectionIndex, setParams.SectionIndices);

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;
            if (setParams.Top.HasValue) pageSetup.TopMargin = setParams.Top.Value;
            if (setParams.Bottom.HasValue) pageSetup.BottomMargin = setParams.Bottom.Value;
            if (setParams.Left.HasValue) pageSetup.LeftMargin = setParams.Left.Value;
            if (setParams.Right.HasValue) pageSetup.RightMargin = setParams.Right.Value;
        }

        MarkModified(context);

        return Success($"Page margins updated for {sectionsToUpdate.Count} section(s)");
    }

    /// <summary>
    ///     Extracts set margins parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set margins parameters.</returns>
    private static SetMarginsParameters ExtractSetMarginsParameters(OperationParameters parameters)
    {
        return new SetMarginsParameters(
            parameters.GetOptional<double?>("top"),
            parameters.GetOptional<double?>("bottom"),
            parameters.GetOptional<double?>("left"),
            parameters.GetOptional<double?>("right"),
            parameters.GetOptional<int?>("sectionIndex"),
            parameters.GetOptional<JsonArray?>("sectionIndices")
        );
    }

    /// <summary>
    ///     Record to hold set margins parameters.
    /// </summary>
    /// <param name="Top">The top margin in points.</param>
    /// <param name="Bottom">The bottom margin in points.</param>
    /// <param name="Left">The left margin in points.</param>
    /// <param name="Right">The right margin in points.</param>
    /// <param name="SectionIndex">The section index to apply margins to.</param>
    /// <param name="SectionIndices">The array of section indices to apply margins to.</param>
    private sealed record SetMarginsParameters(
        double? Top,
        double? Bottom,
        double? Left,
        double? Right,
        int? SectionIndex,
        JsonArray? SectionIndices);
}
