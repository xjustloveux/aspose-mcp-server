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
        var setParams = ExtractSetOrientationParameters(parameters);

        if (string.IsNullOrEmpty(setParams.Orientation))
            throw new ArgumentException("orientation parameter is required for set_orientation operation");

        var doc = context.Document;
        var orientationEnum = string.Equals(setParams.Orientation, "landscape", StringComparison.OrdinalIgnoreCase)
            ? Orientation.Landscape
            : Orientation.Portrait;
        var sectionsToUpdate = WordPageHelper.GetTargetSections(doc, setParams.SectionIndex, setParams.SectionIndices);

        foreach (var idx in sectionsToUpdate)
            doc.Sections[idx].PageSetup.Orientation = orientationEnum;

        MarkModified(context);

        return Success($"Page orientation set to {setParams.Orientation} for {sectionsToUpdate.Count} section(s)");
    }

    /// <summary>
    ///     Extracts set orientation parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set orientation parameters.</returns>
    private static SetOrientationParameters ExtractSetOrientationParameters(OperationParameters parameters)
    {
        return new SetOrientationParameters(
            parameters.GetOptional<string?>("orientation"),
            parameters.GetOptional<int?>("sectionIndex"),
            parameters.GetOptional<JsonArray?>("sectionIndices")
        );
    }

    /// <summary>
    ///     Record to hold set orientation parameters.
    /// </summary>
    /// <param name="Orientation">The page orientation (Portrait or Landscape).</param>
    /// <param name="SectionIndex">The section index to apply orientation to.</param>
    /// <param name="SectionIndices">The array of section indices to apply orientation to.</param>
    private sealed record SetOrientationParameters(string? Orientation, int? SectionIndex, JsonArray? SectionIndices);
}
