using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for clearing tab stops in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ClearTabStopsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "clear_tab_stops";

    /// <summary>
    ///     Clears tab stops from a paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractClearTabStopsParameters(parameters);

        var doc = context.Document;
        var para = WordFormatHelper.GetTargetParagraph(doc, p.ParagraphIndex);

        var count = para.ParagraphFormat.TabStops.Count;
        para.ParagraphFormat.TabStops.Clear();

        MarkModified(context);
        return new SuccessResult { Message = $"Cleared {count} tab stop(s) from paragraph {p.ParagraphIndex}" };
    }

    private static ClearTabStopsParameters ExtractClearTabStopsParameters(OperationParameters parameters)
    {
        return new ClearTabStopsParameters(
            parameters.GetOptional("paragraphIndex", 0));
    }

    private sealed record ClearTabStopsParameters(int ParagraphIndex);
}
