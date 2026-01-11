using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for setting page orientation in Word documents.
/// </summary>
public class SetOrientationWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_orientation";

    /// <summary>
    ///     Sets page orientation (portrait or landscape) for the specified sections.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: orientation (Portrait or Landscape)
    ///     Optional: sectionIndex, sectionIndices
    /// </param>
    /// <returns>Success message with orientation details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var orientation = parameters.GetOptional<string?>("orientation");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var sectionIndices = parameters.GetOptional<JsonArray?>("sectionIndices");

        if (string.IsNullOrEmpty(orientation))
            throw new ArgumentException("orientation parameter is required for set_orientation operation");

        var doc = context.Document;
        var orientationEnum = orientation.ToLower() == "landscape" ? Orientation.Landscape : Orientation.Portrait;
        var sectionsToUpdate = WordPageHelper.GetTargetSections(doc, sectionIndex, sectionIndices);

        foreach (var idx in sectionsToUpdate)
            doc.Sections[idx].PageSetup.Orientation = orientationEnum;

        MarkModified(context);

        return Success($"Page orientation set to {orientation} for {sectionsToUpdate.Count} section(s)");
    }
}
