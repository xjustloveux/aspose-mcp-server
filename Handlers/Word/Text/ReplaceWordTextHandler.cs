using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for replacing text in Word documents.
/// </summary>
public class ReplaceWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "replace";

    /// <summary>
    ///     Replaces text in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: find, replace.
    ///     Optional: useRegex, replaceInFields.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when find or replace parameters are missing.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractReplaceParameters(parameters);

        var doc = context.Document;

        var options = new FindReplaceOptions();
        if (!p.ReplaceInFields)
            options.ReplacingCallback = new FieldSkipReplacingCallback();

        if (p.UseRegex)
            doc.Range.Replace(new Regex(p.Find, RegexOptions.None, TimeSpan.FromSeconds(30)), p.Replace, options);
        else
            doc.Range.Replace(p.Find, p.Replace, options);

        MarkModified(context);

        var result = "Text replaced in document.";
        if (!p.ReplaceInFields)
            result += " Note: Fields (such as hyperlinks) were excluded from replacement.";

        return Success(result);
    }

    private static ReplaceParameters ExtractReplaceParameters(OperationParameters parameters)
    {
        return new ReplaceParameters(
            parameters.GetRequired<string>("find"),
            parameters.GetRequired<string>("replace"),
            parameters.GetOptional("useRegex", false),
            parameters.GetOptional("replaceInFields", false));
    }

    private record ReplaceParameters(
        string Find,
        string Replace,
        bool UseRegex,
        bool ReplaceInFields);
}

/// <summary>
///     Callback to skip field replacement during text replacement operations.
/// </summary>
internal class FieldSkipReplacingCallback : IReplacingCallback
{
    /// <summary>
    ///     Determines whether to replace or skip text replacement based on field context.
    /// </summary>
    /// <param name="args">Replacing arguments containing match information.</param>
    /// <returns>ReplaceAction.Skip if inside a field, ReplaceAction.Replace otherwise.</returns>
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        if (args.MatchNode.GetAncestor(NodeType.FieldStart) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldSeparator) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldEnd) != null)
            return ReplaceAction.Skip;
        return ReplaceAction.Replace;
    }
}
