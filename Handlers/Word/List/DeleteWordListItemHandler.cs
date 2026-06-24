using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for deleting list items from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteWordListItemHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_item";

    /// <summary>
    ///     Deletes a list item at the specified paragraph index.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteListItemParameters(parameters);

        var doc = context.Document;

        var paraToDelete = ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, p.ParagraphIndex))
            .Paragraph;

        var itemText = paraToDelete.GetText().Trim();
        var itemPreview = itemText.Length > 50 ? string.Concat(itemText.AsSpan(0, 50), "...") : itemText;
        var isListItem = paraToDelete.ListFormat.IsListItem;
        var listInfo = isListItem ? " (list item)" : " (regular paragraph)";

        paraToDelete.Remove();
        MarkModified(context);

        var result = $"List item #{p.ParagraphIndex} deleted successfully{listInfo}\n";
        if (!string.IsNullOrEmpty(itemPreview)) result += $"Content preview: {itemPreview}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}";

        return new SuccessResult { Message = result };
    }

    private static DeleteListItemParameters ExtractDeleteListItemParameters(OperationParameters parameters)
    {
        return new DeleteListItemParameters(
            parameters.GetRequired<int>("paragraphIndex"));
    }

    private sealed record DeleteListItemParameters(int ParagraphIndex);
}
