using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for deleting list items from Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteListItemParameters(parameters);

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (p.ParagraphIndex < 0 || p.ParagraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {p.ParagraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (paragraphs[p.ParagraphIndex] is not WordParagraph paraToDelete)
            throw new InvalidOperationException($"Unable to get paragraph at index {p.ParagraphIndex}");

        var itemText = paraToDelete.GetText().Trim();
        var itemPreview = itemText.Length > 50 ? string.Concat(itemText.AsSpan(0, 50), "...") : itemText;
        var isListItem = paraToDelete.ListFormat.IsListItem;
        var listInfo = isListItem ? " (list item)" : " (regular paragraph)";

        paraToDelete.Remove();
        MarkModified(context);

        var result = $"List item #{p.ParagraphIndex} deleted successfully{listInfo}\n";
        if (!string.IsNullOrEmpty(itemPreview)) result += $"Content preview: {itemPreview}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}";

        return Success(result);
    }

    private static DeleteListItemParameters ExtractDeleteListItemParameters(OperationParameters parameters)
    {
        return new DeleteListItemParameters(
            parameters.GetRequired<int>("paragraphIndex"));
    }

    private sealed record DeleteListItemParameters(int ParagraphIndex);
}
