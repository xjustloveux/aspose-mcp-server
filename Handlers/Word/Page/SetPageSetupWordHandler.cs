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
        var setParams = ExtractSetPageSetupParameters(parameters);

        var doc = context.Document;
        var idx = setParams.SectionIndex ?? 0;

        if (idx < 0 || idx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var pageSetup = doc.Sections[idx].PageSetup;
        List<string> changes = [];

        if (setParams.Top.HasValue)
        {
            pageSetup.TopMargin = setParams.Top.Value;
            changes.Add($"Top margin: {setParams.Top.Value}");
        }

        if (setParams.Bottom.HasValue)
        {
            pageSetup.BottomMargin = setParams.Bottom.Value;
            changes.Add($"Bottom margin: {setParams.Bottom.Value}");
        }

        if (setParams.Left.HasValue)
        {
            pageSetup.LeftMargin = setParams.Left.Value;
            changes.Add($"Left margin: {setParams.Left.Value}");
        }

        if (setParams.Right.HasValue)
        {
            pageSetup.RightMargin = setParams.Right.Value;
            changes.Add($"Right margin: {setParams.Right.Value}");
        }

        if (!string.IsNullOrEmpty(setParams.Orientation))
        {
            pageSetup.Orientation =
                string.Equals(setParams.Orientation, "landscape", StringComparison.OrdinalIgnoreCase)
                    ? Orientation.Landscape
                    : Orientation.Portrait;
            changes.Add($"Orientation: {setParams.Orientation}");
        }

        MarkModified(context);

        return $"Page setup updated: {string.Join(", ", changes)}";
    }

    /// <summary>
    ///     Extracts set page setup parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set page setup parameters.</returns>
    private static SetPageSetupParameters ExtractSetPageSetupParameters(OperationParameters parameters)
    {
        return new SetPageSetupParameters(
            parameters.GetOptional<double?>("top"),
            parameters.GetOptional<double?>("bottom"),
            parameters.GetOptional<double?>("left"),
            parameters.GetOptional<double?>("right"),
            parameters.GetOptional<string?>("orientation"),
            parameters.GetOptional<int?>("sectionIndex")
        );
    }

    /// <summary>
    ///     Record to hold set page setup parameters.
    /// </summary>
    /// <param name="Top">The top margin in points.</param>
    /// <param name="Bottom">The bottom margin in points.</param>
    /// <param name="Left">The left margin in points.</param>
    /// <param name="Right">The right margin in points.</param>
    /// <param name="Orientation">The page orientation (Portrait or Landscape).</param>
    /// <param name="SectionIndex">The section index to apply settings to.</param>
    private record SetPageSetupParameters(
        double? Top,
        double? Bottom,
        double? Left,
        double? Right,
        string? Orientation,
        int? SectionIndex);
}
