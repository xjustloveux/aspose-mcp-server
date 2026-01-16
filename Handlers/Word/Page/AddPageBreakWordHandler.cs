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
        var addParams = ExtractAddPageBreakParameters(parameters);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        if (addParams.ParagraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            if (addParams.ParagraphIndex.Value < 0 || addParams.ParagraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");

            builder.MoveTo(paragraphs[addParams.ParagraphIndex.Value]);
            builder.InsertBreak(BreakType.PageBreak);
        }
        else
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
        }

        MarkModified(context);

        var location = addParams.ParagraphIndex.HasValue
            ? $"after paragraph {addParams.ParagraphIndex.Value}"
            : "at document end";
        return Success($"Page break added {location}");
    }

    /// <summary>
    ///     Extracts add page break parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add page break parameters.</returns>
    private static AddPageBreakParameters ExtractAddPageBreakParameters(OperationParameters parameters)
    {
        return new AddPageBreakParameters(
            parameters.GetOptional<int?>("paragraphIndex")
        );
    }

    /// <summary>
    ///     Record to hold add page break parameters.
    /// </summary>
    /// <param name="ParagraphIndex">The 0-based paragraph index to insert page break after.</param>
    private sealed record AddPageBreakParameters(int? ParagraphIndex);
}
