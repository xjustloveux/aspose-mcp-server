using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for clearing tab stops in Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional("paragraphIndex", 0);

        var doc = context.Document;
        var para = WordFormatHelper.GetTargetParagraph(doc, paragraphIndex);

        var count = para.ParagraphFormat.TabStops.Count;
        para.ParagraphFormat.TabStops.Clear();

        MarkModified(context);
        return Success($"Cleared {count} tab stop(s) from paragraph {paragraphIndex}");
    }
}
