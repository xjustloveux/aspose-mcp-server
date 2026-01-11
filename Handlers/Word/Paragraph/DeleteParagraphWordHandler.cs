using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for deleting paragraphs from Word documents.
/// </summary>
public class DeleteParagraphWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a paragraph at the specified index.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");

        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for delete operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        var idx = paragraphIndex.Value;
        if (idx == -1)
        {
            if (paragraphs.Count == 0)
                throw new ArgumentException("Cannot delete paragraph: document has no paragraphs");
            idx = paragraphs.Count - 1;
        }

        if (idx < 0 || idx >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {idx} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}, or -1 for last).");

        var paragraphToDelete = paragraphs[idx] as Aspose.Words.Paragraph;
        if (paragraphToDelete == null)
            throw new InvalidOperationException($"Unable to get paragraph at index {idx}");

        var textPreview = paragraphToDelete.GetText().Trim();
        if (textPreview.Length > 50) textPreview = textPreview.Substring(0, 50) + "...";

        paragraphToDelete.Remove();

        MarkModified(context);

        var result = $"Paragraph #{idx} deleted successfully\n";
        if (!string.IsNullOrEmpty(textPreview)) result += $"Content preview: {textPreview}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}";

        return result;
    }
}
