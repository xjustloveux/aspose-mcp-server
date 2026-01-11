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
        var paragraphIndex = parameters.GetRequired<int>("paragraphIndex");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (paragraphs[paragraphIndex] is not WordParagraph paraToDelete)
            throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex}");

        var itemText = paraToDelete.GetText().Trim();
        var itemPreview = itemText.Length > 50 ? itemText.Substring(0, 50) + "..." : itemText;
        var isListItem = paraToDelete.ListFormat.IsListItem;
        var listInfo = isListItem ? " (list item)" : " (regular paragraph)";

        paraToDelete.Remove();
        MarkModified(context);

        var result = $"List item #{paragraphIndex} deleted successfully{listInfo}\n";
        if (!string.IsNullOrEmpty(itemPreview)) result += $"Content preview: {itemPreview}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}";

        return Success(result);
    }
}
