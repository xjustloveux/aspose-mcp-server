using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for setting multiple page setup options (margins and orientation) in Word documents.
/// </summary>
public class SetPageSetupWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_page_setup";

    /// <summary>
    ///     Sets multiple page setup options (margins and orientation) for a section.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: top, bottom, left, right (margins in points)
    ///     Optional: orientation (Portrait or Landscape)
    ///     Optional: sectionIndex
    /// </param>
    /// <returns>Success message with changes made.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var top = parameters.GetOptional<double?>("top");
        var bottom = parameters.GetOptional<double?>("bottom");
        var left = parameters.GetOptional<double?>("left");
        var right = parameters.GetOptional<double?>("right");
        var orientation = parameters.GetOptional<string?>("orientation");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var idx = sectionIndex ?? 0;

        if (idx < 0 || idx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var pageSetup = doc.Sections[idx].PageSetup;
        List<string> changes = [];

        if (top.HasValue)
        {
            pageSetup.TopMargin = top.Value;
            changes.Add($"Top margin: {top.Value}");
        }

        if (bottom.HasValue)
        {
            pageSetup.BottomMargin = bottom.Value;
            changes.Add($"Bottom margin: {bottom.Value}");
        }

        if (left.HasValue)
        {
            pageSetup.LeftMargin = left.Value;
            changes.Add($"Left margin: {left.Value}");
        }

        if (right.HasValue)
        {
            pageSetup.RightMargin = right.Value;
            changes.Add($"Right margin: {right.Value}");
        }

        if (!string.IsNullOrEmpty(orientation))
        {
            pageSetup.Orientation = orientation.ToLower() == "landscape" ? Orientation.Landscape : Orientation.Portrait;
            changes.Add($"Orientation: {orientation}");
        }

        MarkModified(context);

        return $"Page setup updated: {string.Join(", ", changes)}";
    }
}
