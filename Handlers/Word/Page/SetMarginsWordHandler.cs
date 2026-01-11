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
        var top = parameters.GetOptional<double?>("top");
        var bottom = parameters.GetOptional<double?>("bottom");
        var left = parameters.GetOptional<double?>("left");
        var right = parameters.GetOptional<double?>("right");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var sectionIndices = parameters.GetOptional<JsonArray?>("sectionIndices");

        var doc = context.Document;
        var sectionsToUpdate = WordPageHelper.GetTargetSections(doc, sectionIndex, sectionIndices);

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;
            if (top.HasValue) pageSetup.TopMargin = top.Value;
            if (bottom.HasValue) pageSetup.BottomMargin = bottom.Value;
            if (left.HasValue) pageSetup.LeftMargin = left.Value;
            if (right.HasValue) pageSetup.RightMargin = right.Value;
        }

        MarkModified(context);

        return Success($"Page margins updated for {sectionsToUpdate.Count} section(s)");
    }
}
