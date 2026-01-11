using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Text;

/// <summary>
///     Handler for replacing text in PowerPoint presentations.
/// </summary>
public class ReplacePptTextHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "replace";

    /// <summary>
    ///     Replaces text across all shapes in the presentation.
    ///     Searches text in AutoShapes, GroupShapes (recursive), and Table cells.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: findText, replaceText.
    ///     Optional: matchCase (default: false).
    /// </param>
    /// <returns>Success message with replacement count.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var findText = parameters.GetRequired<string>("findText");
        var replaceText = parameters.GetRequired<string>("replaceText");
        var matchCase = parameters.GetOptional("matchCase", false);

        var presentation = context.Document;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var replacements = 0;

        foreach (var slide in presentation.Slides)
            replacements += PptTextHelper.ProcessShapesForReplace(slide.Shapes, findText, replaceText, comparison);

        MarkModified(context);

        return Success($"Replaced '{findText}' with '{replaceText}' ({replacements} occurrences).");
    }
}
