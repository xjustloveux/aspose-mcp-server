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
        var replaceParams = ExtractReplaceParameters(parameters);

        var presentation = context.Document;
        var comparison = replaceParams.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var replacements = 0;

        foreach (var slide in presentation.Slides)
            replacements += PptTextHelper.ProcessShapesForReplace(slide.Shapes, replaceParams.FindText,
                replaceParams.ReplaceText, comparison);

        MarkModified(context);

        return Success(
            $"Replaced '{replaceParams.FindText}' with '{replaceParams.ReplaceText}' ({replacements} occurrences).");
    }

    /// <summary>
    ///     Extracts replace parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted replace parameters.</returns>
    private static ReplaceParameters ExtractReplaceParameters(OperationParameters parameters)
    {
        return new ReplaceParameters(
            parameters.GetRequired<string>("findText"),
            parameters.GetRequired<string>("replaceText"),
            parameters.GetOptional("matchCase", false)
        );
    }

    /// <summary>
    ///     Record for holding replace text parameters.
    /// </summary>
    /// <param name="FindText">The text to find.</param>
    /// <param name="ReplaceText">The text to replace with.</param>
    /// <param name="MatchCase">Whether to match case.</param>
    private record ReplaceParameters(string FindText, string ReplaceText, bool MatchCase);
}
