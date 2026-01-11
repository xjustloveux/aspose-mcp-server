using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for adding a page break at a specified paragraph or at document end in Word documents.
/// </summary>
public class AddPageBreakWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_page_break";

    /// <summary>
    ///     Adds a page break at the specified paragraph or at document end.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: paragraphIndex (0-based paragraph index to insert page break after)
    /// </param>
    /// <returns>Success message with page break location.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");

            builder.MoveTo(paragraphs[paragraphIndex.Value]);
            builder.InsertBreak(BreakType.PageBreak);
        }
        else
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
        }

        MarkModified(context);

        var location = paragraphIndex.HasValue ? $"after paragraph {paragraphIndex.Value}" : "at document end";
        return Success($"Page break added {location}");
    }
}
